#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent
#Include IMEv2.ahk

UIA_TextPatternId := 10014
CLSID_CUIAutomation8 := "{E22AD333-B25F-460C-83D0-0581107395C9}"
IID_IUIAutomation := "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}"
MAX_ROMAJI_LENGTH := 128
LOG_FILE := A_ScriptDir "\boiled-egg-debug.log"

; Wait tuning. State-driven waits remain; fixed waits are kept only where
; Excel/IME needs time to consume a key event.
EXCEL_F2_TIMEOUT_MS := 300
EXCEL_DELETE_SETTLE_MS := 30
EXCEL_IME_SETTLE_MS := 25
CLIPBOARD_TIMEOUT_MS := 300
UIA_BOUNDARY_TIMEOUT_MS := 180
EXCEL_BOUNDARY_TIMEOUT_MS := 300

g_UIAObject := 0
g_pUIA := 0
g_ImeOnAllowed := false
g_ImePolicySuspended := false

InitializeUIAutomation()
SetImeOff()
SetTimer(EnforceImePolicy, 30)
OnExit(BoiledEgg_OnExit)

^j::BoiledEgg()
^h::SendEvent("{Backspace}")
^#!v::PastePlainText()
^m::HandleCtrlM()
^o::ToggleIme()
SC07B::SetImeOff()  ; 無変換
SC079::SetImeOn()   ; 変換
^Esc::ExitApp()

PastePlainText()
{
    ; ホットキー修飾キーが貼り付けへ混ざらないよう、解放を待つ。
    KeyWait("Ctrl")
    KeyWait("Alt")
    KeyWait("LWin")
    KeyWait("RWin")

    savedClipboard := ClipboardAll()

    try {
        plainText := A_Clipboard

        if (plainText = "") {
            ShowBoiledEggStatus("クリップボードにテキストがありません")
            return
        }

        ; テキスト形式だけのクリップボードへ一時的に置き換える。
        A_Clipboard := plainText
        ClipWait(0.3)

        SendEvent("^v")

        ; 対象アプリがクリップボードを読み終えるまで少し待つ。
        Sleep(100)
    } finally {
        A_Clipboard := savedClipboard
        savedClipboard := ""
    }
}

SetImeOff()
{
    global g_ImeOnAllowed
    g_ImeOnAllowed := false
    IME_SET(0, "A")
}

SetImeOn()
{
    global g_ImeOnAllowed
    g_ImeOnAllowed := true
    IME_SET(1, "A")
    IME_SetConvMode(25, "A")
}

ToggleIme()
{
    if (IME_GET("A") = 0)
        SetImeOn()
    else
        SetImeOff()
}

EnforceImePolicy()
{
    global g_ImeOnAllowed, g_ImePolicySuspended
    if (!g_ImePolicySuspended && !g_ImeOnAllowed && IME_GET("A") != 0)
        IME_SET(0, "A")
}

HandleCtrlM()
{
    if (IME_GET("A") = 0) {
        SendEvent("{Enter}")
        return
    }
    SetImeOff()
}

BoiledEgg()
{
    global MAX_ROMAJI_LENGTH, g_ImePolicySuspended, g_ImeOnAllowed
    global EXCEL_DELETE_SETTLE_MS, EXCEL_IME_SETTLE_MS

    KeyWait("Ctrl")
    KeyWait("j")

    ; Ctrl+J処理中はIME抑止タイマーを完全に停止する。
    ; フラグだけでなくタイマー自体を止め、IME ON直後の競合を防ぐ。
    SetTimer(EnforceImePolicy, 0)
    g_ImePolicySuspended := true

    runId := FormatTime(, "yyyyMMdd-HHmmss") "-" A_TickCount
    operationStarted := A_TickCount
    targetHwnd := WinExist("A")
    processName := SafeProcessName()
    sourceText := ""
    selectionActive := false
    clipboardScan := false
    clipboardBoundaryHit := false
    excelClipboardScan := false

    LogStage(runId, "start", "process=" processName " hwnd=" targetHwnd)

    try {
        excelState := PrepareExcelForBoiledEgg(runId)
        LogStage(
            runId,
            "excel-state",
            "mode=" excelState.Mode " elapsed=" excelState.Elapsed
        )

        if !excelState.Ok {
            ShowBoiledEggStatus(excelState.Reason)
            return
        }

        Loop MAX_ROMAJI_LENGTH {
            if (WinExist("A") != targetHwnd) {
                if (selectionActive)
                    CollapseSelectionForMethod(clipboardScan)
                LogStage(runId, "window-changed")
                ShowBoiledEggStatus("対象ウィンドウが変わりました")
                return
            }

            SendEvent("+{Left}")
            selectionActive := true
            ; Excelは次の取得処理が完了待ちを兼ねるため、固定待機を置かず
            ; 一度タイムスライスを譲る。その他はUIA反映用に10ms待つ。
            Sleep(IsExcelWindow() ? 0 : 10)

            selected := GetSelectedText()
            LogStage(
                runId,
                "scan",
                "n=" A_Index
                " ok=" selected.Ok
                " method=" selected.Method
                " text=" QuoteLog(selected.Text)
                " reason=" selected.Reason
            )

            if !selected.Ok {
                if (selectionActive)
                    CollapseSelectionForMethod(clipboardScan)
                ShowBoiledEggStatus("選択範囲を取得できませんでした: " selected.Reason)
                return
            }

            candidate := selected.Text

            if (selected.Method = "clipboard") {
                if IsExcelWindow() {
                    excelClipboardScan := true
                    ; Excelの編集欄ではCtrl+Insert後も選択が維持され、
                    ; コピー結果は累積文字列になるためUIA経路と同様に扱う。
                    if IsAllowedScanText(candidate) {
                        ; Excelではセル先頭などでShift+Leftを追加しても
                        ; 選択範囲が広がらず、同じ文字列が返り続けることがある。
                        ; 同一結果を左端境界として扱い、現在の選択を採用する。
                        if (candidate = sourceText && sourceText != "") {
                            LogStage(
                                runId,
                                "boundary-stable",
                                "method=excel-clipboard"
                                " text=" QuoteLog(sourceText)
                                " n=" A_Index
                            )
                            selectionActive := true
                            break
                        }

                        sourceText := candidate
                        selectionActive := true
                        continue
                    }

                    LogStage(
                        runId,
                        "boundary-detected",
                        "method=excel-clipboard candidate=" QuoteLog(candidate)
                        " expected=" QuoteLog(sourceText)
                        " firstCode=" FirstCodePoint(candidate)
                    )

                    if !TrimAndVerifyBoundaryClipboard(sourceText, runId) {
                        selectionActive := false
                        ShowBoiledEggStatus("Excelの境界文字を選択範囲から外せませんでした")
                        return
                    }

                    LogStage(runId, "boundary-trim-ok", "method=excel-clipboard text=" QuoteLog(sourceText))
                    break
                }

                ; Emacsなど、コピー後に選択が解除されるアプリ用。
                clipboardScan := true
                selectionActive := false

                if IsAllowedScanChar(candidate) {
                    sourceText := candidate sourceText
                    continue
                }

                clipboardBoundaryHit := true
                LogStage(runId, "boundary", "method=clipboard text=" QuoteLog(candidate))
                break
            }

            ; UIAでは累積した選択文字列を使う。
            if IsAllowedScanText(candidate) {
                sourceText := candidate
                continue
            }

            LogStage(
                runId,
                "boundary-detected",
                "method=uia candidate=" QuoteLog(candidate)
                " expected=" QuoteLog(sourceText)
                " firstCode=" FirstCodePoint(candidate)
            )

            if !TrimAndVerifyBoundaryUIA(sourceText, runId) {
                selectionActive := false
                ShowBoiledEggStatus("境界文字を選択範囲から外せませんでした")
                return
            }

            LogStage(runId, "boundary-trim-ok", "text=" QuoteLog(sourceText))
            break
        }

        if (sourceText = "") {
            if (selectionActive)
                CollapseSelectionForMethod(clipboardScan)
            LogStage(runId, "no-source-text")
            ShowBoiledEggStatus("キャレット直前に対象文字がありません")
            return
        }

        if (clipboardScan) {
            ; 1文字ずつ左へ走査した位置から、対象全体を選び直す。
            if (clipboardBoundaryHit)
                SendEvent("{Right}")
            SendEvent("+{Right " StrLen(sourceText) "}")
            selectionActive := true
            Sleep(0)
            LogStage(runId, "reselect-done", "length=" StrLen(sourceText))
        }

        if (excelClipboardScan) {
            ; F2で編集モードへ入ったExcelでは、IME入力が現在の選択を
            ; 置換しない場合がある。選択したローマ字を先に削除し、
            ; 同じ位置へIME入力する。
            LogStage(
                runId,
                "excel-source-delete",
                "text=" QuoteLog(sourceText)
            )
            ; Excelでは選択文字列の削除にDeleteを使う。
            SendEvent("{Delete}")
            Sleep(EXCEL_DELETE_SETTLE_MS)
            selectionActive := false
            LogStage(
                runId,
                "excel-source-delete-done",
                "imeBefore=" IME_GET("A")
            )
        }

        ; IMEへはローマ字を小文字で送る。先頭大文字に必要なShiftが
        ; Excel/IMEの入力モードへ影響することを避ける。
        imeInputText := StrLower(sourceText)

        LogStage(
            runId,
            "ime-on-request",
            "source=" QuoteLog(sourceText)
            " input=" QuoteLog(imeInputText)
        )
        SetImeOn()
        LogStage(runId, "ime-set-called", "allowed=" g_ImeOnAllowed " state=" IME_GET("A"))

        if !WaitForImeState(1, 300, 5) {
            if (selectionActive)
                CollapseSelectionForMethod(clipboardScan)
            LogStage(runId, "ime-on-fail", "state=" IME_GET("A"))
            ShowBoiledEggStatus("IME ONを確認できませんでした")
            return
        }

        LogStage(runId, "ime-on-ok", "state=" IME_GET("A"))
        IME_SetConvMode(25, "A")
        Sleep(excelClipboardScan ? EXCEL_IME_SETTLE_MS : 10)
        LogStage(
            runId,
            "conv-mode",
            "mode=" IME_GetConvMode("A")
            " process=" SafeProcessName()
        )

        if (clipboardScan) {
            ; Emacsでは文字入力がリージョンを置換しないため明示的に削除。
            savedClipboard := ClipboardAll()
            try {
                LogStage(runId, "source-delete", "text=" QuoteLog(sourceText))
                SendEvent("+{Delete}")
                Sleep(10)
                LogStage(runId, "source-delete-done")
            } finally {
                A_Clipboard := savedClipboard
            }
            selectionActive := false
        }

        if (excelClipboardScan)
            SetKeyDelay(25, 10)
        else
            SetKeyDelay(35, 25)

        LogStage(
            runId,
            "reinput",
            "text=" QuoteLog(imeInputText)
            " original=" QuoteLog(sourceText)
            " ime=" IME_GET("A")
            " allowed=" g_ImeOnAllowed
            " excelClipboard=" excelClipboardScan
            " sendMethod=SendEvent"
        )
        ; IMEへ実際のキー入力として渡すため、ExcelでもSendEventを使う。
        ; SendTextは文字をリテラル送信するため、IME変換を開始しない場合がある。
        SendEvent(imeInputText)
        selectionActive := false
        LogStage(
            runId,
            "success",
            "text=" QuoteLog(sourceText)
            " totalElapsed=" (A_TickCount - operationStarted)
        )
        ShowBoiledEggStatus(sourceText " をIMEへ再入力しました")
    } catch Error as err {
        LogStage(runId, "exception", "message=" QuoteLog(err.Message) " line=" err.Line)
        if (selectionActive)
            CollapseSelectionForMethod(clipboardScan)
        ShowBoiledEggStatus("エラー: " err.Message)
    } finally {
        ; 再変換処理が完全に終了してから監視を再開する。
        g_ImePolicySuspended := false
        SetTimer(EnforceImePolicy, 30)
    }
}

PrepareExcelForBoiledEgg(runId)
{
    global EXCEL_F2_TIMEOUT_MS
    started := A_TickCount

    if !IsExcelWindow() {
        return {Ok:true, Mode:"not-excel", Reason:"", Elapsed:0}
    }

    ; Excelのセル内編集ではEXCEL6クラスの子ウィンドウが現れる。
    ; UIAが使えない環境でも、これで既編集中か判定できる。
    if IsExcelNativeEditMode() {
        return {
            Ok:true,
            Mode:"already-editing-native",
            Reason:"",
            Elapsed:A_TickCount - started
        }
    }

    ; 数式バーなど、UIA TextPatternが使える編集状態も確認する。
    probe := GetSelectedTextUIA()
    if probe.Ok {
        return {
            Ok:true,
            Mode:"already-editing-uia",
            Reason:"",
            Elapsed:A_TickCount - started
        }
    }

    ; 編集ウィンドウもTextPatternもないため、セル選択状態と判断してF2。
    LogStage(runId, "excel-f2-request", "uia=" probe.Reason)
    SendEvent("{F2}")

    ; EXCEL6が現れるまで短時間待つ。現れなくても従来の
    ; クリップボードフォールバックを利用して処理を続ける。
    while (A_TickCount - started < EXCEL_F2_TIMEOUT_MS) {
        Sleep(5)
        if IsExcelNativeEditMode() {
            LogStage(runId, "excel-f2-result", "nativeEdit=1")
            return {
                Ok:true,
                Mode:"entered-editing-native",
                Reason:"",
                Elapsed:A_TickCount - started
            }
        }
    }

    probeAfterF2 := GetSelectedTextUIA()
    LogStage(
        runId,
        "excel-f2-result",
        "nativeEdit=0 uiaOk=" probeAfterF2.Ok
        " reason=" probeAfterF2.Reason
    )

    return {
        Ok:true,
        Mode:probeAfterF2.Ok ? "entered-editing-uia" : "entered-editing-fallback",
        Reason:"",
        Elapsed:A_TickCount - started
    }
}

IsExcelNativeEditMode()
{
    if !IsExcelWindow()
        return false

    try {
        rootHwnd := WinExist("A")
        return FindDescendantWindowByClass(rootHwnd, "EXCEL6") != 0
    } catch {
        return false
    }
}

FindDescendantWindowByClass(parentHwnd, targetClass)
{
    childHwnd := 0

    Loop {
        childHwnd := DllCall(
            "user32\FindWindowExW",
            "Ptr", parentHwnd,
            "Ptr", childHwnd,
            "Ptr", 0,
            "Ptr", 0,
            "Ptr"
        )

        if !childHwnd
            break

        classBuffer := Buffer(256 * 2, 0)
        length := DllCall(
            "user32\GetClassNameW",
            "Ptr", childHwnd,
            "Ptr", classBuffer.Ptr,
            "Int", 256,
            "Int"
        )

        if (length > 0) {
            className := StrGet(classBuffer, length, "UTF-16")
            if (className = targetClass)
                return childHwnd
        }

        descendant := FindDescendantWindowByClass(childHwnd, targetClass)
        if descendant
            return descendant
    }

    return 0
}

IsExcelWindow()
{
    try
        return StrLower(WinGetProcessName("A")) = "excel.exe"
    catch
        return false
}

TrimAndVerifyBoundaryUIA(expectedText, runId, timeoutMs := 0)
{
    global UIA_BOUNDARY_TIMEOUT_MS
    if (timeoutMs <= 0)
        timeoutMs := UIA_BOUNDARY_TIMEOUT_MS

    ; 「。Soukamo」から「Soukamo」へ1文字だけ縮める。
    SendEvent("+{Right}")

    started := A_TickCount
    lastText := ""
    lastReason := "not checked"

    while (A_TickCount - started < timeoutMs) {
        Sleep(5)
        selected := GetSelectedTextUIA()
        lastReason := selected.Reason

        if selected.Ok {
            lastText := selected.Text
            normalizedText := NormalizeUIASelectionText(lastText)

            ; NotepadのTextPatternは、行頭をまたいだ選択を縮めた後も
            ; 選択範囲の端にCR/LFを付けて返すことがある。
            ; 実際の対象文字列の比較では前後の改行だけを除外する。
            if (normalizedText = expectedText) {
                LogStage(
                    runId,
                    "boundary-verify",
                    "ok=1 elapsed=" (A_TickCount - started)
                    " expected=" QuoteLog(expectedText)
                    " raw=" QuoteLog(lastText)
                    " normalized=" QuoteLog(normalizedText)
                )
                return true
            }
        }
    }

    LogStage(
        runId,
        "boundary-verify",
        "ok=0 elapsed=" (A_TickCount - started)
        " expected=" QuoteLog(expectedText)
        " raw=" QuoteLog(lastText)
        " normalized=" QuoteLog(NormalizeUIASelectionText(lastText))
        " reason=" lastReason
    )

    ; 不正な後方選択を右端へ畳む。
    SendEvent("{Right}")
    return false
}


TrimAndVerifyBoundaryClipboard(expectedText, runId, timeoutMs := 0)
{
    global EXCEL_BOUNDARY_TIMEOUT_MS
    if (timeoutMs <= 0)
        timeoutMs := EXCEL_BOUNDARY_TIMEOUT_MS
    SendEvent("+{Right}")
    started := A_TickCount
    lastText := ""
    lastReason := "not checked"

    while (A_TickCount - started < timeoutMs) {
        Sleep(0)
        selected := GetSelectedTextByClipboardShortcut("Excel boundary verification")
        lastReason := selected.Reason

        if selected.Ok {
            lastText := NormalizeUIASelectionText(selected.Text)
            if (lastText = expectedText) {
                LogStage(
                    runId,
                    "boundary-verify",
                    "ok=1 method=excel-clipboard"
                    " elapsed=" (A_TickCount - started)
                    " expected=" QuoteLog(expectedText)
                    " actual=" QuoteLog(lastText)
                )
                return true
            }
        }
    }

    LogStage(
        runId,
        "boundary-verify",
        "ok=0 method=excel-clipboard"
        " expected=" QuoteLog(expectedText)
        " actual=" QuoteLog(lastText)
        " reason=" lastReason
    )
    SendEvent("{Right}")
    return false
}

NormalizeUIASelectionText(text)
{
    ; UIAが選択範囲の端へ付加するCR/LFだけを比較対象から除外する。
    ; 文字列内部の改行は変更しない。
    return RegExReplace(text, "^[\r\n]+|[\r\n]+$", "")
}

IsAllowedScanChar(text)
{
    ; 走査対象は1文字の半角ASCIIだけ。
    if (StrLen(text) != 1)
        return false

    code := Ord(text)
    if (code > 0x7F)
        return false

    return RegExMatch(text, "^[A-Za-z0-9.,!?-]$") != 0
}

IsAllowedScanText(text)
{
    if (text = "")
        return false

    Loop Parse text {
        if !IsAllowedScanChar(A_LoopField)
            return false
    }

    return true
}

GetSelectedText()
{
    selected := GetSelectedTextUIA()
    if selected.Ok
        return selected
    return GetSelectedTextByClipboardShortcut(selected.Reason)
}

GetSelectedTextByClipboardShortcut(uiaFailureReason := "")
{
    global CLIPBOARD_TIMEOUT_MS
    savedClipboard := ClipboardAll()

    try {
        ; 空にした後の番号を基準にする。
        A_Clipboard := ""
        beforeSequence := DllCall("user32\GetClipboardSequenceNumber", "UInt")
        SendEvent("^{Insert}")
        started := A_TickCount

        while (A_TickCount - started < CLIPBOARD_TIMEOUT_MS) {
            currentSequence := DllCall("user32\GetClipboardSequenceNumber", "UInt")
            if (currentSequence != beforeSequence) {
                try {
                    text := A_Clipboard
                    if (text = "") {
                        Sleep(2)
                        continue
                    }
                    return {
                        Ok:true,
                        Text:text,
                        Reason:"clipboard fallback; UIA=" uiaFailureReason,
                        Method:"clipboard"
                    }
                } catch Error as err {
                    return {
                        Ok:false,
                        Text:"",
                        Reason:"clipboard read failed: " err.Message,
                        Method:"clipboard"
                    }
                }
            }
            Sleep(2)
        }

        return {
            Ok:false,
            Text:"",
            Reason:"clipboard fallback timed out; UIA=" uiaFailureReason,
            Method:"clipboard"
        }
    } finally {
        A_Clipboard := savedClipboard
    }
}

InitializeUIAutomation()
{
    global g_UIAObject, g_pUIA, CLSID_CUIAutomation8, IID_IUIAutomation
    g_UIAObject := ComObject(CLSID_CUIAutomation8, IID_IUIAutomation)
    g_pUIA := ComObjValue(g_UIAObject)
}

GetSelectedTextUIA()
{
    global g_pUIA, UIA_TextPatternId

    pElement := 0
    pPattern := 0
    pRanges := 0
    pRange := 0
    pBSTR := 0
    result := {Ok:false, Text:"", Reason:"not checked", Method:"uia"}

    try {
        vtUIA := NumGet(g_pUIA, 0, "Ptr")
        hr := DllCall(
            NumGet(vtUIA, 8 * A_PtrSize, "Ptr"),
            "Ptr", g_pUIA,
            "Ptr*", &pElement,
            "Int"
        )
        if (HResultFailed(hr) || !pElement) {
            result.Reason := "focused element unavailable"
            return result
        }

        vtElement := NumGet(pElement, 0, "Ptr")
        hr := DllCall(
            NumGet(vtElement, 16 * A_PtrSize, "Ptr"),
            "Ptr", pElement,
            "Int", UIA_TextPatternId,
            "Ptr*", &pPattern,
            "Int"
        )
        if (HResultFailed(hr) || !pPattern) {
            result.Reason := "TextPattern unavailable"
            return result
        }

        vtPattern := NumGet(pPattern, 0, "Ptr")
        hr := DllCall(
            NumGet(vtPattern, 5 * A_PtrSize, "Ptr"),
            "Ptr", pPattern,
            "Ptr*", &pRanges,
            "Int"
        )
        if (HResultFailed(hr) || !pRanges) {
            result.Reason := "selection unavailable"
            return result
        }

        vtRanges := NumGet(pRanges, 0, "Ptr")
        count := 0
        hr := DllCall(
            NumGet(vtRanges, 3 * A_PtrSize, "Ptr"),
            "Ptr", pRanges,
            "Int*", &count,
            "Int"
        )
        if (HResultFailed(hr) || count != 1) {
            result.Reason := "selection range count is not one"
            return result
        }

        hr := DllCall(
            NumGet(vtRanges, 4 * A_PtrSize, "Ptr"),
            "Ptr", pRanges,
            "Int", 0,
            "Ptr*", &pRange,
            "Int"
        )
        if (HResultFailed(hr) || !pRange) {
            result.Reason := "selection range unavailable"
            return result
        }

        vtRange := NumGet(pRange, 0, "Ptr")
        hr := DllCall(
            NumGet(vtRange, 12 * A_PtrSize, "Ptr"),
            "Ptr", pRange,
            "Int", -1,
            "Ptr*", &pBSTR,
            "Int"
        )
        if HResultFailed(hr) {
            result.Reason := "GetText failed"
            return result
        }

        result.Text := pBSTR ? StrGet(pBSTR, "UTF-16") : ""
        result.Ok := true
        result.Reason := "retrieved"
        return result
    } catch Error as err {
        result.Reason := "UIA exception: " err.Message
        return result
    } finally {
        if pBSTR
            DllCall("oleaut32\SysFreeString", "Ptr", pBSTR)
        if pRange
            ComReleaseRaw(pRange)
        if pRanges
            ComReleaseRaw(pRanges)
        if pPattern
            ComReleaseRaw(pPattern)
        if pElement
            ComReleaseRaw(pElement)
    }
}

WaitForImeState(expectedState, timeoutMs, pollMs)
{
    started := A_TickCount
    while (A_TickCount - started < timeoutMs) {
        if (IME_GET("A") = expectedState)
            return true
        Sleep(pollMs)
    }
    return false
}

CollapseSelectionForMethod(clipboardScan)
{
    if clipboardScan
        SendEvent("{Left}")
    else
        SendEvent("{Right}")
}

LogStage(runId, stage, detail := "")
{
    global LOG_FILE
    try {
        line := FormatTime(, "yyyy-MM-dd HH:mm:ss")
            . " [" runId "] " stage
        if (detail != "")
            line .= " | " detail
        FileAppend(line "`n", LOG_FILE, "UTF-8")
    }
}

FirstCodePoint(text)
{
    if (text = "")
        return "none"
    return Format("U+{:04X}", Ord(SubStr(text, 1, 1)))
}

QuoteLog(text)
{
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "`n", "\n")
    text := StrReplace(text, "`t", "\t")
    return Chr(34) text Chr(34)
}

SafeProcessName()
{
    try
        return WinGetProcessName("A")
    catch
        return "unknown"
}

ShowBoiledEggStatus(message)
{
    ToolTip("boiled-egg: " message)
    SetTimer(() => ToolTip(), -1500)
}

BoiledEgg_OnExit(reason, code)
{
    global g_UIAObject, g_pUIA
    g_pUIA := 0
    g_UIAObject := 0
}

HResultFailed(hr)
{
    return hr < 0
}

ComReleaseRaw(p)
{
    if !p
        return 0
    vt := NumGet(p, 0, "Ptr")
    return DllCall(
        NumGet(vt, 2 * A_PtrSize, "Ptr"),
        "Ptr", p,
        "UInt"
    )
}
