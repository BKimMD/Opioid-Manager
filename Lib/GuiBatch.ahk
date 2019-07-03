GuiBatch(PtInfo, BatchTitle)
{
	Gui, Destroy
	static PauseButton

	UdsDate := PtInfo[13]
	NaloxDate :=P tInfo[19]

	FormatTime, UdsDate, %UdsDate%, ShortDate
	FormatTime, NaloxDate, %NaloxDate%, ShortDate

	Gui, 99:New , +AlwaysOnTop +Resize, %BatchTitle%
	Gui, 99:Add, Text,, % PtInfo[1]", "PtInfo[2]", SSN: "PtInfo[3]"`n"PtInfo[37]" || Last Rx:"PtInfo[36]"`nLast UDS:" UdsDate " || Last Naloxone:" NaloxDate

	IsPaused := false
	BttnLabel := 0
	Gui, 99:Add, Button, w100 vPauseButton Default, Pause
	;% BttnLabel?"Paused":"Running"
	Gui, 99:Show, y0
	return

	ButtonPause:
	if IsPaused
	{
		Pause off
		IsPaused := false
		GuiControl,, PauseButton, Pause
	}
	else
		SetTimer, Pause, 10
	return

	Pause:
	SetTimer, Pause, off
	IsPaused := true
	GuiControl,, PauseButton, Unpause
	Pause, on
	return
	
	99GuiEscape:
	99GuiClose:
		ExitApp
}