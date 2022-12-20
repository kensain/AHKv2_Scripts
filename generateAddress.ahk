#SingleInstance Force

#x::ExitApp
#HotIf WinActive("ahk_exe EXCEL.EXE")
F1::generateAddress()

A_TrayMenu.Delete()         ; Delete default tray menu items.
A_TrayMenu.Add("Справка", ObjBindMethod(about))     ; Add an 'About' menu with a brief description of how the script works.
; Persistent

/**
 * Creates a word document based on the value from the currently selected cell.
 */
generateAddress() {
    excel := ComObjActive("Excel.Application")
    word := ComObjActive("Word.Application")
    
    customerAddress := excel.ActiveCell.Value       ; The value of the currenntly selected cell will be the customer's address
    customerName := excel.Activecell.Offset(0,-1).Value     ; With an offset we shift 0 rows and 1 column left to select a cell with the customer's name
    word.Documents.Add      ; Create new word document
    word.ActiveDocument.PageSetup.Orientation := 1      ; Change orientation to the 'landscape'
    word.ActiveDocument.Paragraphs(1).SpaceAfter := 2       ; Adjust the spacing between the paragraphs so it is not so big
    word.Selection.Font.Size := 14      ; Set font size
    word.Selection.Font.Bold := 1       ; Set bold style
    word.Selection.ParagraphFormat.Alignment := 2       ; Set text alignment to 'Right'
    word.activedocument.Content.InsertAfter("`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n")       ; Add some new lines so the text would be printed on the right bottom side
    word.activedocument.Content.InsertAfter(customerName "`n")      ; Insert customer name from the offset cell
    word.activedocument.Content.InsertAfter(RegExReplace(customerAddress, ",\s", ",`r`n"))      ; Insert the customer address

    ; The two lines below better exist
    word.Visible := 1       ; Show the word window
    word.Activate       ; Activate the word window
}

/**
 * Add some description to the script. For example, what is it about and what hotkeys are required.
 */
about(*) {
    MsgBox("Для работы должны быть запущены программы Word и Excel.`n`nКликаем на ячейку с адресом в Excel, нажимаем F1(работает только из Excel), создаётся лист А4 с адресом в нужной ориентации. Готово, можно отправлять на печать!`n`nДля выхода из программы нужно нажать Win+X", "Справка")
}