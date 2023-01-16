#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Run, C:\Program Files (x86)\ESP\ESPPLUS\Syzer\SystemSyzer.exe
sleep, 5000

SetControlDelay -1
ControlClick, x182 y234, System Syzer V4.4 ; click on the Flow/Pressure Drop tab
sleep, 200

FilePath := "C:\Users\RIchardC\Documents\Local Revit Copies\Brandon Avenue\PipeMiniSchedule.xlsx"
oExcel := ComObjCreate("Excel.Application")
oExcel.DisplayAlerts := False
oWorkBook := oExcel.Workbooks.Open(FilePath)
oExcel.Visible := False
sleep, 200

sheetName := "CHW C"

xlUp := -4162
BottomRow := oWorkBook.Worksheets(sheetName).Rows.Count
LastRow := oWorkBook.Worksheets(sheetName).Range("B" BottomRow).End(xlUp).Row ; method to get last row of sheet

for _, i in range(1, LastRow) { ; loop thru all rows in sheet
	ControlGet, cList, List, , WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad14, System Syzer V4.4
	Loop, Parse, cList, `n
	{
		; if InStr(A_LoopField, oWorkBook.Worksheets(sheetName).Range("B" i)) {
		worksheetDiameter := oWorkBook.Worksheets(sheetName).Range("B" i).Value
		if (A_LoopField == worksheetDiameter) {
			
			if(worksheetDiameter == "3 in" or worksheetDiameter == "4 in" or worksheetDiameter == "6 in") { ; set the pipe material to steel if diameter is 3, 4, or 6 inches
				ControlGet, cList, List, , WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad13, System Syzer V4.4
				Loop, Parse, cList, `n
				{
					if InStr(A_LoopField, "Steel Pipe") {
						Control, ChooseString, %A_LoopField%, WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad13, System Syzer V4.4
					}
				}
			} else {  ; set the pipe material to copper otherwise
				ControlGet, cList, List, , WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad13, System Syzer V4.4
				Loop, Parse, cList, `n
				{
					if InStr(A_LoopField, "Type L Copper") {
						Control, ChooseString, %A_LoopField%, WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad13, System Syzer V4.4
					}
				}
			}
			sleep, 200
			
			Control, ChooseString, %A_LoopField%, WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad14, System Syzer V4.4 ; set the pipe size
			
			ControlSend, WindowsForms10.EDIT.app.0.141b42a_r8_ad144, {Ctrl down}a{Ctrl up}, System Syzer V4.4
			flowRate := oExcel.Worksheets(sheetName).Range("D" i).Value
			; MsgBox %flowRate%
			ControlSend, WindowsForms10.EDIT.app.0.141b42a_r8_ad144, %flowRate%, System Syzer V4.4 ; input the flow rate
			sleep, 200

			ControlClick, x300 y234, System Syzer V4.4 ; click on the Length/Pressure Drop
			sleep, 200

			ControlGetText, fLoss, WindowsForms10.EDIT.app.0.141b42a_r8_ad154, System Syzer V4.4 ; store the friction loss in fLoss
			oWorkBook.Worksheets(sheetName).Range("E" i).Value := fLoss

			;~ ControlGetText, hLoss, WindowsForms10.EDIT.app.0.141b42a_r8_ad155, System Syzer V4.4 ; store the friction loss in fLoss
			;~ oWorkBook.Worksheets(sheetName).Range("F" i).Value := hLoss

			ControlClick, x182 y234, System Syzer V4.4 ; click on the Flow/Pressure Drop tab
			sleep, 200
			break
		}
		; sleep, 200
	}
}
oWorkBook.Save()
oWorkBook.Close()
MsgBox Finished Export to PipeMiniSchedule


range(startx, endx, stepsize := 1) {
	stepsize := stepsize * (startx < endx ? 1 : -1)
	range_a := Array()
	Loop {
		range_a.Push(startx)
		startx += stepsize
	} Until ((stepsize > 0) ? (startx >= endx) : (startx <= endx))
	range_a.Push(startx)
	return range_a
}