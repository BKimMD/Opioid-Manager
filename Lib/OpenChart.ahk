OpenChart(PatientInfo, BatchTitle)
{
	WinWaitActive, Patient Selection ahk_class TfrmPtSel
	GuiBatch(PatientInfo, BatchTitle)
;	FirstLetter := Substr(PatientInfo[1], 1, 1)
	ControlFocus, TORComboEdit3, Patient Selection ahk_class TfrmPtSel
	ControlSend, TORComboEdit3, {BackSpace 15},Patient Selection ahk_class TfrmPtSel
	ControlSend, TORComboEdit3, % PatientInfo[3],Patient Selection ahk_class TfrmPtSel

	Loop {
		ControlGet, ListVar, List, Count 1, TORListBox2, Patient Selection ahk_class TfrmPtSel
		If InStr(ListVar, PatientInfo[1])
			break
		sleep 50
	}
	control, check,, TButton9
	
	Sleep 500
	Start := A_TickCount
	while (A_TickCount-Start <= 1000)
	{
		IfWinNotActive, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe
			Start := A_TickCount
		while (A_Cursor = AppStarting) or (A_Cursor = Wait)
			sleep 500
	}

	while (A_Cursor = AppStarting) or (A_Cursor = Wait)
		sleep 500
	
;	msgbox % PatientInfo[1]","PatientInfo[2]" "PatientInfo[3]
;	send {BackSpace 5}
;	ControlSend, TORComboEdit3, % PatientInfo[1]","PatientInfo[2], Patient Selection ahk_class TfrmPtSel
;	ControlClick, TButton9
	;Patient Selection ahk_class TfrmPtSel

}