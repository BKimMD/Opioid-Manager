;Main file for UDSBatch.exe
;For a given provider’s long-term opioid patients, creates UDS orders and mail out letters if overdue
;Bryan Kim, MD
;https://github.com/BKimMD/Opioid-Manager
;BryanKimMD@gmail.com

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force
#InstallKeybdHook       ;Installs needed hook for hotstring/hotkeys
#UseHook                ;Enables needed hook for hotstring/hotkeys (can also use $ in front of hotkey)
SetControlDelay, 200
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
MonthsSinceDone :=5
DateDiff := MonthsSinceDone * 30 ;5 Months x 30 days
DueDate := A_Now
DueDate += -DateDiff, days

url := "https://reports.vssc.med.va.gov/ReportServer/Pages/ReportViewer.aspx?%2fPC%2fAlmanac%2fPAIN_ProviderWEB&rs:Command=Render"
run, %url%
sleep 2000

msgbox,
(
This program will automatically find patients missing a UDS in the last %MonthsSinceDone% months, and create an order and letter.

1. Select patient population (team, opioid population, opioid group)
2. Save list as a CSV File
3. Open the saved CSV file on next prompt
)

FileSelectFile, BatchFile,,,Select CSV List to Order UDS and Send Note, Comma Separated Value File (*.CSV)
ListOfNames := CSVtoArray(BatchFile)

;ListOfNames := CSVtoArray("PAIN_ProviderWEB.csv")

#IfWinActive, ahk_exe CPRSChart.exe ;CPRS 2.0
WinActivate, ahk_exe CPRSChart.exe
IfWinNotExist, Patient Selection ahk_class TfrmPtSel
	send !fn
WinWaitActive, Patient Selection ahk_class TfrmPtSel

For k, v in ListOfNames
{	
	if (k>1 and v[13] < DueDate and v[32] = "" ) ;removes header, checks past duedate, and skips palliative pts
	{
;		msgbox % v[1]v[2]v[3]v[32]
		OpenChart(v, "Batch UDS Program")
		if !UdsAlreadyOrdered()
		{

		OrderUDS()
;		send ^n		
		LastDone := v[13]
;		msgbox % "Preformat="LastDone
		FormatTime, LastDone, %LastDone%, ShortDate
;		msgbox % "Formatted="LastDone
		Message =`n`tI am writing to inform you that per VA policy, completion of a Urine Toxicology screening is required twice a year or as directed by your primary care provider.`n`n`tAccording to our records you are due for a 6 month urine drug screen which needs to be completed within this refill and no further refills will be issued until the Urine Toxicology Screening is completed. There is no need to schedule an appointment with your primary care provider at this time. Lab orders for the urine drug screen have been placed for you, please report to lab at your earliest convenience to have these tests completed.`n`n`t**Last Urine Tox Screen completed: %LastDone%**`n`n`tIf you have any further questions or concerns, please call the phone number at the top of this letter.
		LetterWriter(Message)
		Msgbox Review and Sign orders, and go to new patient when ready
		}
	}

;	WinWaitActive,Review / Sign Changes ahk_class TfrmReview,,30
;	WinWaitClose, Review / Sign Changes ahk_class TfrmReview,,30
;	if ErrorLevel
;		break
}
Gui, Destroy
#IfWinActive
Msgbox UDS script completed, exiting
ExitApp

UdsAlreadyOrdered()
{
	UDSOrdered := false
	WinWaitActive, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe
	send ^r
	while !CPRSPage("Reports Page")
		sleep 500
	FindReportPage("Lab Status", "L")

/*
	ControlSend,TORTreeView1,{home}, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe, Reports Page
	Loop
	{
		ControlSend,TORTreeView1,L{space}, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe, Reports Page
		sleep 50
;		ControlClick, TORTreeView1, ahk_class TfrmFrame, Reports Page
		ControlGetText, TreeLabel, TCaptionListView1, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe, Reports Page
		sleep 50

		If InStr(TreeLabel, "Lab Status")
			break

		if A_Index > 10
		{
			msgbox Lab Status not found, exiting
			break
		}
	}
*/
	Control, Check,, TRadioButton7, ahk_class TfrmFrame ahk_exe CPRSChart.exe, Reports Page ;select labs within last month
;	WinWaitActive, ahk_class TfrmFrame ahk_exe CPRSChart.exe, Reports Page
	
		Loop 
		{
			ControlGetText, LabTxt, TRichEdit1
			sleep 50
			ControlGetText, LabTxt2, TRichEdit1
		} Until LabTxt = LabTxt2
		
	If InStr(LabTxt, "OPIATES           ROUTINE Requested")
	{
		UDSOrdered := true
		MsgBox Previous UDS order still active, skipping to next patient
		send !fn
	}
;	msgbox % UDSOrdered
return UDSOrdered
}

OrderUDS()
{
	WinWaitActive, VistA CPRS ahk_exe CPRSChart.exe
	CoordMode, Mouse, Screen
	MouseMove, 0, 0
	send ^o
	while !CPRSPage("Orders Page")
		sleep 500
;	WinWaitActive, ahk_class TfrmFrame, Orders Page
;	ControlFocus, TORAlignButton1, ahk_class TfrmFrame, Orders Page
	ControlSend, TORListBox1, L, ahk_class TfrmFrame, Orders Page
;	Send {tab}{PGDN 2}{Up 10}{enter}
	WinWaitActive, ahk_class TfrmOMNavA, Labs...
	Send, {right 2}{down 4}{enter}
	WinWaitActive, ahk_class TfrmOMNavA, TODAY'S OUTPATIENT LABS
	Send, {right 2}{down 19}{enter}
	WinWaitClose, URINE TOXICOLOGY ORDER SET ahk_class TfrmOMNavA
	WinWaitActive, ahk_class TfrmOMNavA, TODAY'S OUTPATIENT LABS
	Control, check,,TORAlignButton1
	WinWaitClose, ahk_class TfrmOMNavA, TODAY'S OUTPATIENT LABS
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