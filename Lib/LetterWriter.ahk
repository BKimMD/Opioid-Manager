LetterWriter(Message)
{
	SetControlDelay, 200
	SetKeyDelay, 1, 1
	
	WinWaitActive, VistA CPRS ahk_exe CPRSChart.exe
	Send ^n
	sleep 50
	
	while (A_Cursor = AppStarting) or (A_Cursor = Wait)
		sleep 500
	
	Send +^n
	WinWaitActive, Progress Note Properties ahk_class TfrmNoteProperties,,20
	send FOLLOW UP LETTER{enter}
	WinWaitActive, ahk_class TfrmTemplateDialog,,20

	control, check,,TCPRSDialogParentCheckBox5,A ;HWND FOR LOCATION
	control, check,,TCPRSDialogParentCheckBox8,A ;HWND FOR RMR
;	control, check,,TCPRSDialogParentCheckBox18,A ;HWND for Labs
	control, check,,TButton4,A
	
	WinWaitActive, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes

	while (A_Cursor = AppStarting) or (A_Cursor = Wait)
		sleep 500
	
	ControlGetText, LetterString, TRichEdit3, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
	LetterString .= Message
;	msgbox % LetterString
	ControlSetText, TRichEdit3, %LetterString%, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
	;var .= var1 .= var2 .= var3 https://autohotkey.com/board/topic/29596-combining-variables/
	ControlFocus, TRichEdit3, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
	ControlSend, TRichEdit3, {PGDN}, ahk_class TfrmFrame ahk_exe CPRSChart.exe, frmNotes
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