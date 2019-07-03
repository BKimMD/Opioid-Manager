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
;SetControlDelay, 100
SetTitleMatchMode, 2
DetectHiddenText, Off
DetectHiddenWindows, On
SetKeyDelay, 1, 1, 1   ;Sets the delay that will occur after each keystroke sent by Send and ControlSend. SetKeyDelay [, Delay, PressDuration, Play]1 

IfWinNotExist, ahk_exe CPRSChart.exe
{
	MsgBox, Open CPRS first then start this program
	ExitApp
}

GroupAdd, ComplexCloseGroup, Restricted Record ahk_class #32770 ahk_exe CPRSChart.exe ;Closes Restricted Record
GroupAdd, SimpleCloseGroup, Patient Record Flags ahk_class TfrmFlags ahk_exe CPRSChart.exe ;Closes Patient Record Flags
GroupAdd, SimpleCloseGroup, Patient Lookup Messages ahk_class TfrmPtSelMsg ahk_exe CPRSChart.exe
;GroupAdd, CloseGroup, ahk_class TfrmSignon

;ahk_group CloseGroup

SetTimer, ClosePopup, 50
SetTimer, MainThread, -1
return

MainThread:
MonthsSinceDone :=5
DateDiff := MonthsSinceDone * 30 ;5 Months x 30 days
DueDate := A_Now
DueDate += -DateDiff, days

Global IsWinClosedGlobal := ""
Global IsWinOpenGlobal := ""

/*
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
*/

;ListOfNames := CSVtoArray("U:\My Documents\PAIN_ProviderWEB BK.csv")
;ListOfNames := CSVtoArray("U:\My Documents\Mcbryde 6-19 PAIN_ProviderWEB (1).csv")
ListOfNames := CSVtoArray("U:\My Documents\PAIN_ProviderWEB SCRUBBED.csv")



#IfWinActive, ahk_exe CPRSChart.exe ;CPRS 2.0
WinActivate, ahk_exe CPRSChart.exe

For k, v in ListOfNames
{	
	if (k>1 and v[13] < DueDate and v[32] = "" ) ;removes header, checks past duedate, and skips palliative pts
	{
;		msgbox % v[1]v[2]v[3]v[32]

		OpenChart(v, "Batch UDS Program")
		if !UdsAlreadyOrdered()
		{
			OrderUDS()

			LastDone := v[13]
	;		msgbox % "Preformat="LastDone
			FormatTime, LastDone, %LastDone%, ShortDate
	;		msgbox % "Formatted="LastDone
			Message =`tI am writing to inform you that per VA policy, completion of a Urine Toxicology screening is required twice a year or as directed by your primary care provider.`n`n`tAccording to our records you are due for a 6 month urine drug screen which needs to be completed within this refill and no further refills will be issued until the Urine Toxicology Screening is completed. There is no need to schedule an appointment with your primary care provider at this time. Lab orders for the urine drug screen have been placed for you, please report to lab at your earliest convenience to have these tests completed.`n`n`t**Last Urine Tox Screen completed: %LastDone%**`n`n`tIf you have any further questions or concerns, please call the phone number at the top of this letter.
			LetterWriter(Message)
;			Msgbox Review and Sign orders. Click OK to go to next patient when ready
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
	GotoCPRSPage("Report")

	;SendTorTreeView(GoalTT, SendKeys, WinTitle, WinText:="", GoalOffset:=0, TButtonToggle:="", PageShortcutKey:="")
	SendTorTreeView("Available Reports", "{LEFT 4}L{SPACE}", "VistA CPRS in use by ahk_class TfrmFrame ahk_exe CPRSChart.exe")
;	FindReportPage("Lab Status", "{left 4}L{space}")

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
	Control, Check,, TRadioButton7 ;select labs within last month
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
	}
;	msgbox % UDSOrdered
return UDSOrdered
}

OrderUDS()
{	
	WinWaitActive, VistA CPRS ahk_exe CPRSChart.exe
	GotoCPRSPage("Orders")
	CoordMode, Mouse, Screen
	MouseMove, 0, 0
	
;	WinWaitActive, ahk_class TfrmFrame, Orders Page
;	ControlFocus, TORListBox1, ahk_class TfrmFrame, Orders Page,, Consults Page
;	send L
;


	ControlSend, TORListBox1, L, ahk_class TfrmFrame ahk_exe CPRSChart.exe, Orders Page ;Brings up Lab page

/*
	WinWait, Order Menu ahk_class TfrmOMNavA ahk_exe CPRSChart.exe, Consult Orders,3
	WinActivate
	send {escape}
	send ^o
*/

;	WinActivate, Order Menu ahk_class TfrmOMNavA ahk_exe CPRSChart.exe, Labs`. ;Activates Lab page

;;	Send {tab}{PGDN 2}{Up 10}{enter}
					
;	WinWaitActive
	WinWait, Order Menu ahk_class TfrmOMNavA ahk_exe CPRSChart.exe, Labs`.

	ControlSend, TCaptionStringGrid1,{right 2}{down 4}{enter}

	WinWait, Order Menu ahk_class TfrmOMNavA, TODAY'S OUTPATIENT LABS
	ControlSend, TCaptionStringGrid1, {right 2}{down 19}{enter}
	
	IsWinClosedGlobal := "URINE TOXICOLOGY ORDER SET ahk_class TfrmOMSet"
;	BusyPause("URINE TOXICOLOGY ORDER SET ahk_class TfrmOMSet")
	
	WinWaitActive, Order Menu ahk_class TfrmOMNavA, TODAY'S OUTPATIENT LABS
	Control, check,, TORAlignButton1
	send {enter}
	WinWaitClose

}

;-------------------------------------------------------------------------

ClosePopup:	;Runs In Background and closes Popups
	If !WinExist("CPRS Assist") ;Disable concurrent popup blocks if CPRS Assist is active
		{
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
				return
			}
		}

	IfWinExist, ahk_group SimpleCloseGroup ;Closes one button dialogue boxes (see top of script)
	{
		WinActivate
		send {enter}
		WinWaitClose
		return
	}

	IfWinExist, ahk_group ComplexCloseGroup ;Closes multi-button dialogue boxes (see top of script)
	{
	
		WinActivate

		Loop
		{
			ButtonNumber := A_Index
;			msgbox % ButtonNumber
			controlgettext, ButtonText, Button%ButtonNumber%
			If A_Index >4 
			{
				ButtonNumber := 0
				msgbox Can't find Ok Button
				Break
			}
		} until (InStr(ButtonText,"ok") or InStr(ButtonText,"yes") or InStr(ButtonText,"close"))
		if ButtonNumber
			control, check,, Button%ButtonNumber%
		WinWaitClose
		return
	}
	
	BusyPause()

return