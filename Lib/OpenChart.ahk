OpenChart(PatientInfo, BatchTitle)
{
	If !WinExist("Patient Selection ahk_class TfrmPtSel ahk_exe CPRSChart.exe")
		WinMenuSelectItem, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe, , File, Select ;Selects Next Patient

	WinWaitActive, Patient Selection ahk_class TfrmPtSel ahk_exe CPRSChart.exe
	GuiBatch(PatientInfo, BatchTitle)
;	WinActivate
;	FirstLetter := Substr(PatientInfo[1], 1, 1)
;	ControlFocus, TORComboEdit3
	ControlSetText, TORComboEdit3 ;Clears out existing text
;	ControlSend, TORComboEdit3, {BackSpace 15}
	ControlSend, TORComboEdit3, % PatientInfo[3]

	Loop {
		ControlGet, ListVar, List, Count 1, TORListBox2
		If InStr(ListVar, PatientInfo[1])
			break
		sleep 50
	}
	control, check,, TButton9
	sleep 50
	WinWaitClose, Similar Patients ahk_class TfrmDupPts ahk_exe CPRSChart.exe ;Pause for similar names check

	IsWinOpenGlobal :="VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe"
	
;	BusyPause(,"VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe")
	
;	msgbox % PatientInfo[1]","PatientInfo[2]" "PatientInfo[3]
;	send {BackSpace 5}
;	ControlSend, TORComboEdit3, % PatientInfo[1]","PatientInfo[2], Patient Selection ahk_class TfrmPtSel
;	ControlClick, TButton9
	;Patient Selection ahk_class TfrmPtSel

}