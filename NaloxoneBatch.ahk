;Main file for NaloxoneBatch.exe  
;For a given provider’s long-term opioid patients, creates Nalxone orders and mail out letters if overdue
;Bryan Kim, MD
;https://github.com/BKimMD/Opioid-Manager
;BryanKimMD@gmail.com

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force
#InstallKeybdHook       ;Installs Needed hook for hotstring/hotkeys
#UseHook                ;Enables Needed hook for hotstring/hotkeys (can also use $ in front of hotkey)
SetControlDelay, 300
SetTitleMatchMode, 2
DetectHiddenText, Off
SetKeyDelay, 1,1

IfWinNotExist, ahk_exe CPRSChart.exe
{
	MsgBox, Open CPRS first then start this program
	ExitApp
}

SetTimer, MainThread, -1
SetTimer, ClosePopup, 50
return

MainThread:
MonthsSinceDone:=13
DaysUntilOverdue := MonthsSinceDone*30 ;13 months x 30 days
DueDate := A_Now
DueDate += -DaysUntilOverdue, days

url := "https://reports.vssc.med.va.gov/ReportServer/Pages/ReportViewer.aspx?%2fPC%2fAlmanac%2fPAIN_ProviderWEB&rs:Command=Render"
run, %url%
sleep 2000

msgbox,
(
This program will automatically find patients without Naloxone in the last %MonthsSinceDone% months, and create an order and letter.

1. Select patient population
2. Save list as a CSV File
3. Open saved CSV file
)

;msgbox %A_ScriptDir%
FileSelectFile, NaloxeneBatchFile,,,Select CSV List to Send Naloxone Kits and Send Note, Comma Separated Value File (*.CSV)
ListOfNames := CSVtoArray(NaloxeneBatchFile)

;ListOfNames := CSVtoArray("PAIN_ProviderWEB.csv")
;FileRead, NaloxeneBatchFile, U:\My Documents\Test UDS and Naloxone.csv
;msgbox % NaloxeneBatchFile

#IfWinActive, ahk_exe CPRSChart.exe ;CPRS 2.0
WinActivate, ahk_exe CPRSChart.exe
IfWinNotExist, Patient Selection ahk_class TfrmPtSel
	send !fn
WinWaitActive, Patient Selection ahk_class TfrmPtSel

For k, v in ListOfNames
{
;	msgbox % "processing " v[1]v[2]" SSN:"v[3]" Last Date:"v[19]" DueDate "DueDate
	if (k>1 and (v[19] < DueDate or v[19] = "000000")) ;removes header and checks past duedate or includes non-naloxone pts (padded)
	{
;		clipboard := v[19]
;		msgbox % "INCLUDED " v[1]v[2]" SSN:"v[3]" Last Date:"v[19]" DueDate "DueDate
		
		OpenChart(v, "Batch Naloxone Program")
		OrderNaloxone()
		
		Message =
		(
Because of the rising opiate epidemic where patients on opiates have overdosed and died from taking opiates, we are taking precautions to prevent overdoses in our veterans on chronic opiates by prescribing NALOXONE, a medication used in emergency cases to reverse opiate overdoses.

Classes are available to veterans, friends and family at:

A4-235 (4th floor Building A)
Rocky Mountain Regional VAMC
1700 North Wheeling St
Aurora, CO 80045

On third Thursdays of the month at 3:30-4:30PM 
Walk-ins welcome, no appointment needed.

For more information, please call (720) 723-7418.
)
		LetterWriter(Message)
	}

;	WinWaitActive,Review / Sign Changes ahk_class TfrmReview,,30
;	WinWaitClose, Review / Sign Changes ahk_class TfrmReview,,30
;	if ErrorLevel
;		break
}
Gui, Destroy
#IfWinActive
MsgBox Naloxone script complete, exiting
ExitApp

OrderNaloxone()
{
	WinWaitActive, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe
	send ^m
	while !CPRSPage("Medications Page")
		sleep 500
;	WinWaitActive, ahk_exe CPRSChart.exe, Outpatient Medications
	send !an
	WinWaitActive, Outpatient Medications ahk_class TfrmODMeds
;	ControlFocus, TEdit1, Outpatient Medications ahk_class TfrmODMeds
;	ControlSetText, TEdit1, NALOXONE RESCUE SOLN
	ControlSend, TEdit1, NALOXONE RESCUE SOLN
;	ControlSend, TEdit1, NALOXONE RESCUE SOLN, Outpatient Medications ahk_class TfrmODMeds
	Loop {
	ControlGetText, MedText, TEdit1
	If InStr(MedText, "NALOXONE RESCUE SOLN,SPRAY,NASAL ")
		break
	sleep 50
	}

/*	ControlGet, Value, Enabled,, Tbutton2, Outpatient Medications ahk_class TfrmODMeds)
	while !Value{
		sleep 50
		msgbox % value
		ControlGet, Value, Enabled,, Tbutton2, Outpatient Medications ahk_class TfrmODMeds)
	}	
	*/
	; RESCUE SOLN,SPRAY,NASAL
	;send NALOXONE RESCUE SOLN,SPRAY,NASALFOLLOW UP LETTER
	

;	ControlGet, ListVar, List, , TCaptionListView1, Outpatient Medications ahk_class TfrmODMeds
;	ControlGet, ListVar, List, , TORListBox2, Patient Selection ahk_class TfrmPtSel
	;Outpatient Medications ahk_class TfrmODMeds
;	msgbox hi
;	msgbox % ListVar
;	msgbox % ErrorLevel

;	control, check,,TButton1,A

;	sleep 500
;	control, check,, TButton2
	send {enter}

	WinWaitActive, Outpatient Medications ahk_class TfrmODMeds, Patient Instruction
	ControlSetText, TCaptionEdit3, 30
	ControlSetText, TCaptionEdit2, 1
	ControlSetText, TCaptionEdit1, 1
;	sleep 500
	send {down}
	control, check,,TButton2,A
;	ControlFocus, TButton2, Outpatient Medications ahk_class TfrmODMeds
;	send {enter}
	winwaitclose, Outpatient Medications ahk_class TfrmODMeds, Patient Instruction
	WinWaitActive, Outpatient Medications ahk_class TfrmODMeds, Quick Orders
	control, check,,TButton2,A
}

;-------------------------------------------------------------------------

ClosePopup:	;Runs In Background

	IfWinExist, Location for Current Activities ahk_class TfrmEncounter ahk_exe CPRSChart.exe, Encounter Location, Provider ;CLOSES LOCATION AND FILLS IN 00
	{
		WinActivate
		send +{tab}{right 2}{tab}
		send 00{enter}
		WinWaitClose
	}
	
	IfWinExist, Order Checking ahk_class TfrmOCAccept ahk_exe CPRSChart.exe, Drug Interaction Monograph ;CLOSES NALOXONE/OPIATE DUPLICATE NOTICE
	{
		WinActivate, Order Checking ahk_class TfrmOCAccept ahk_exe CPRSChart.exe, Drug Interaction Monograph
		ControlGetText, OrderCheckMsg, TRichEdit1
		DeleteNonVAMedCheck:=""
		if InStr(OrderCheckMsg, "Duplicate opioid medications:  [1] NALOXONE RESCUE")
			control, check,, TButton3, Order Checking ahk_class TfrmOCAccept ahk_exe CPRSChart.exe, Drug Interaction Monograph

		if InStr(OrderCheckMsg, "Order Checks could not be done for Drug: NO NON-VA MEDS REPORTED, please complete a manual check for Drug Interactions and Duplicate Therapy.")
			control, check,, TButton3, Order Checking ahk_class TfrmOCAccept ahk_exe CPRSChart.exe, Drug Interaction Monograph

		WinWaitClose, Order Checking ahk_class TfrmOCAccept ahk_exe CPRSChart.exe, Drug Interaction Monograph
	}

	while (A_Cursor = AppStarting) or (A_Cursor = Wait)
	sleep 500

return