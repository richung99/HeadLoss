#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Gui, Add, DropDownList, vList1, CHWS|CHWR|HWS|HWR
Gui, Add, Button, default, OK
Gui, Show,, Choose System
return

GuiClose:
ButtonOK:
Gui, Submit
if(List1 = "CHWS") {
	calcLoss(1)
} else if (List1 = "CHWR"){
	calcLoss(2)	
} else if (List1 = "HWS"){
	calcLoss(3)
} else {
	calcLoss(4)
}
ExitApp

calcLoss(sheetNumber) {
	
	Run, C:\Program Files (x86)\ESP\ESPPLUS\Syzer\SystemSyzer.exe
	sleep, 5000
	
	SetControlDelay -1
	ControlClick, x182 y234, System Syzer V4.4 ; click on the Flow/Pressure Drop tab
	sleep, 500

	ControlGet, cList, List, , WindowsForms10.COMBOBOX.app.0.141b42a_r7_ad13, System Syzer V4.4
	Loop, Parse, cList, `n
	{
		if InStr(A_LoopField, "Type L Copper") {
			Control, ChooseString, %A_LoopField%, WindowsForms10.COMBOBOX.app.0.141b42a_r7_ad13, System Syzer V4.4 ; set the pipe material to copper
		}
		sleep, 500
	}
	FilePath := "C:\Users\RIchardC\Documents\Personal Projects\Dynamo Scripts\PipeCalculator\PipeMiniSchedule.xlsx"
	oExcel := ComObjCreate("Excel.Application")
	oExcel.DisplayAlerts := False
	oWorkBook := oExcel.Workbooks.Open(FilePath)
	oExcel.Visible := False
	sleep, 500

	xlUp := -4162
	BottomRow := oWorkBook.Sheets(sheetNumber).Rows.Count
	LastRow := oWorkBook.Sheets(sheetNumber).Range("B" BottomRow).End(xlUp).Row ; method to get last row of sheet

	for _, i in range(1, LastRow) { ; loop thru all rows in sheet
		ControlGet, cList, List, , WindowsForms10.COMBOBOX.app.0.141b42a_r7_ad14, System Syzer V4.4
		Loop, Parse, cList, `n
		{
			if InStr(A_LoopField, oWorkBook.Sheets(sheetNumber).Range("A" i)) {
				Control, ChooseString, %A_LoopField%, WindowsForms10.COMBOBOX.app.0.141b42a_r7_ad14, System Syzer V4.4 ; set the pipe size
				
				ControlSend, WindowsForms10.EDIT.app.0.141b42a_r7_ad144, {Ctrl down}a{Ctrl up}, System Syzer V4.4
				flowRate := oExcel.Sheets(sheetNumber).Range("C" i).Value
				; MsgBox %flowRate%
				ControlSend, WindowsForms10.EDIT.app.0.141b42a_r7_ad144, %flowRate%, System Syzer V4.4 ; input the flow rate

				ControlClick, x300 y234, System Syzer V4.4 ; click on the Length/Pressure Drop
				sleep, 500

				ControlGetText, fLoss, WindowsForms10.EDIT.app.0.141b42a_r7_ad154, System Syzer V4.4 ; store the friction loss in fLoss
				oWorkBook.Sheets(sheetNumber).Range("D" i).Value := fLoss

				ControlGetText, hLoss, WindowsForms10.EDIT.app.0.141b42a_r7_ad155, System Syzer V4.4 ; store the friction loss in fLoss
				oWorkBook.Sheets(sheetNumber).Range("E" i).Value := hLoss

				ControlClick, x182 y234, System Syzer V4.4 ; click on the Flow/Pressure Drop tab
				sleep, 500
				break
			}
			sleep, 500
		}
	}
	oWorkBook.Save()
	oWorkBook.Close()
	MsgBox Finished Export to PipeMiniSchedule
}

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