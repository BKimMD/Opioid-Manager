SendTorTreeView(GoalTT, SendKeys, WinTitle, WinText:="", GoalOffset:=0, TButtonToggle:="") ;Finds correct TorTree and Sendskeys in WinTitle/WinText
;Example: SendTorTreeView("Templates", "{PgUp}{right}{right}{space}", "VistA CPRS in use by ahk_class TfrmFrame ahk_exe CPRSChart.exe",,,"TBitBtn3")
;Or		  SendTorTreeView("Available Reports", "{LEFT 4}L{SPACE}", "VistA CPRS in use by ahk_class TfrmFrame ahk_exe CPRSChart.exe","Reports Page")
{
	if TButtonToggle
	{
		control, check,, %TButtonToggle%, %WinTitle%, %WinText%
		control, check,, %TButtonToggle%, %WinTitle%, %WinText%
	}

/*
	If WinText
	{
		Loop 
		{
			if A_Index > 10 
			{
				msgbox CPRS Page Timed out
				return 2
			}
			sleep 500	
		} until CPRSPage(WinText)
	}
*/	
	Loop 
	{
		TTNumber:=A_Index
		ControlGetText, TestVariable, TORTreeView%TTNumber%, %WinTitle%, %WinText%
		If A_Index > 8
		{
			msgbox Unable to find Correct TorTree
			return 2
		}
	} until TestVariable = GoalTT


	TTNumber:=TTNumber+GoalOffset
;	msgbox TORTreeView%TTNumber%
	ControlSend, TORTreeView%TTNumber%, %SendKeys%,  %WinTitle%, %WinText%
	return TTNumber
}