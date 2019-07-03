LetterWriter(Message)
{
	SetControlDelay, 200
	SetKeyDelay, 1, 1
	
	WinWaitActive, VistA CPRS ahk_exe CPRSChart.exe
	
	GotoCPRSPage("frmNotes")

	Send +^n
	WinWaitActive, Progress Note Properties ahk_class TfrmNoteProperties,,20
	send FOLLOW UP LETTER{enter}
	WinWaitActive, ahk_class TfrmTemplateDialog,,20

	control, check,,TCPRSDialogParentCheckBox5,A ;HWND FOR LOCATION
	control, check,,TCPRSDialogParentCheckBox8,A ;HWND FOR RMR
;	control, check,,TCPRSDialogParentCheckBox18,A ;HWND for Labs
	control, check,,TButton4,A

/*
	Start := A_TickCount
	while (A_TickCount-Start <= 500)
	{
		IfWinNotActive, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
		{
			Start := A_TickCount
			sleep 500
		}
		while (A_Cursor = AppStarting) or (A_Cursor = Wait)
		{
			Start := A_TickCount
			sleep 500
		}
	}
*/

	ControlFocus, TRichEdit3, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
	ControlGetText, LetterString, TRichEdit3, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
;	Message .= {PGDN}
	LetterString .= Message
;	msgbox % LetterString

/*
	Start := A_TickCount
	while (A_TickCount-Start <= 500)
	{
		IfWinNotActive, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
		{
			Start := A_TickCount
			sleep 500
		}
		while (A_Cursor = AppStarting) or (A_Cursor = Wait)
		{
			Start := A_TickCount
			sleep 500
		}
	}

*/
	ControlSetText, TRichEdit3, %LetterString%, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
	;var .= var1 .= var2 .= var3 https://autohotkey.com/board/topic/29596-combining-variables/

;	ControlSend, TRichEdit3, {PGDN}, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
	;SetControlDelay, 0

/*
	ClipSaved := ClipboardAll ; save the entire Clipboard to the variable ClipSaved
	Clipboard := ""           ; empty the Clipboard (start off empty to allow ClipWait to detect when the text has arrived)
	Clipboard := Message              ; copy this text:

	ClipWait, 10         ; wait max. 10 seconds for the Clipboard to contain data. 
	ClipError := ErrorLevel

	if (!ClipError)         ; If NOT ErrorLevel, ClipWait found data on the Clipboard
		PostMessage, 0x302, , , TRichEdit3, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
;		ControlSend, TRichEdit3, ^v, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes ; paste text
	Sleep, 200
	Clipboard := ClipSaved   ; restore original Clipboard
	ClipWait, 10         ; wait max. 10 seconds for the Clipboard to contain data. 
	ClipSaved =              ; Free the memory in case the Clipboard was very large.
*/
}