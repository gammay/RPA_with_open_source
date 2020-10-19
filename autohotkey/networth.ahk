; Use pause to stop running script
Hotkey, Pause, TheEnd

TheEnd(params*) {
    MsgBox, Exitting
    ExitApp
}

; Function to get tock quote
GetStockQuote(code)
{
    ; Start chrome
    Run chrome.exe --start-maximized, ,max
    Sleep, 3000

    ; Go to finance site
    Send, https://finance.yahoo.com/quote/%code%{ENTER}
    Sleep, 5000

    ; Get quote
    MouseMove, 100,930
    Click, 2
    Sleep, 2000
    Sendinput, ^{c}
    Sleep, 2000
    clipwait, 1,1
    Sleep, 2000
    quote := Clipboard

    ; Close chrome
    Send !{f4} 
    Sleep, 2000

    return %quote%
}

; Main
; Open excel with holdings data
Run EXCEL.EXE D:\holdings.xlsx, ,
; Wait for excel to open - a better way is to wait on event
Sleep, 10000
; Get handle to exel app
X1 := ComObjActive("Excel.Application")

networth := 0

FormatTime, CurrentDateTime,, yyyy-MM-dd HH-mm-ss

code := X1.Range("B2").Value
quote := GetStockQuote(code)
; MsgBox, %CurrentDateTime% %code% %quote%
X1.Range("D2").Value := CurrentDateTime
X1.Range("E2").Value := quote

code := X1.Range("B3").Value
quote := GetStockQuote(code)
; MsgBox, %CurrentDateTime% %code% %quote%
X1.Range("D3").Value := CurrentDateTime
X1.Range("E3").Value := quote

code := X1.Range("B4").Value
quote := GetStockQuote(code)
; MsgBox, %CurrentDateTime% %code% %quote%
X1.Range("D4").Value := CurrentDateTime
X1.Range("E4").Value := quote

; X1.Workbooks.iActiveWorkbook.Save()

return


