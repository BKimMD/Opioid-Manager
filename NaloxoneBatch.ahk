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
;SetControlDelay, 300
SetTitleMatchMode, 2
DetectHiddenText, Off
DetectHiddenWindows, On
SetKeyDelay,,1

IfWinNotExist, ahk_exe CPRSChart.exe
{
	MsgBox, Open CPRS first then start this program
	ExitApp
}

GroupAdd, ComplexCloseGroup, Restricted Record ahk_class #32770 ahk_exe CPRSChart.exe ;Restricted Record
GroupAdd, ComplexCloseGroup, Order Checking ahk_class TfrmOCAccept ahk_exe CPRSChart.exe, Drug Interaction Monograph ;Naloxone interaction
GroupAdd, SimpleCloseGroup, Patient Record Flags ahk_class TfrmFlags ahk_exe CPRSChart.exe ;Patient Record Flags
GroupAdd, SimpleCloseGroup, Patient Lookup Messages ahk_class TfrmPtSelMsg ahk_exe CPRSChart.exe ;Self closing lookup message
GroupAdd, SimpleCloseGroup, Order Checks ahk_class #32770 ahk_exe CPRSChart.exe ;Creatnine Check

SetTimer, ClosePopup, 50
SetTimer, MainThread, -1
return

MainThread:
MonthsSinceDone:=13
DaysUntilOverdue := MonthsSinceDone*30 ;13 months x 30 days
DueDate := A_Now
DueDate += -DaysUntilOverdue, days

Global IsWinClosedGlobal := ""
Global IsWinOpenGlobal := ""

/*
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
*/

;ListOfNames := CSVtoArray("U:\My Documents\PAIN_ProviderWEB BK.csv")
;ListOfNames := CSVtoArray("U:\My Documents\Mcbryde 6-19 PAIN_ProviderWEB (1).csv")
ListOfNames := CSVtoArray("U:\My Documents\PAIN_ProviderWEB SCRUBBED.csv")

#IfWinActive, ahk_exe CPRSChart.exe ;CPRS 2.0
WinActivate, ahk_exe CPRSChart.exe

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
Because of the rising epidemic of long-term opiate patients possibly overdosing and dying, we are taking precautions by prescribing NALOXONE, a medication used in emergency cases to reverse opiate overdoses. NALOXONE does not have many risks and can quickly stop an opioid overdose, possibly save lives.

You are at a higher risk for an overdose if:
1. You drink alcohol.
2. You smoke.
3. You have certain medical conditions, such as COPD, asthma, sleep apnea, and kidney or liver disease.
4. You take certain medications, such as benzodiazepines (alprazolam, clonazepam, diazepam, lorazepam) or sleep medications (eszopiclone, zaleplon, zolpidem). 
5. You take buprenorphine (Suboxone) or methadone.
6. You use heroin or other opioids without a prescription from your doctor.
For your safety, tell your pharmacists and doctors about all your medical conditions and medications you take. 
	
To identify an overdose: 
1. CHECK to see if the person is vomiting or in a deep sleep and is hard to wake up.
2. LISTEN for slow or no heartbeat, slow, strange or no breathing, or gurgling or snoring noises.
3. LOOK for blue lips, fingernails, or skin.
4. TOUCH for clammy or sweaty skin. 

If you think a person is overdosing:
1. Check for a response by shaking them, saying their name, or rubbing the center of their chest bone with your knuckles.
2. If the person does not respond, give naloxone and call 911.
3. If the person is not breathing, start CPR (rescue breathing to provide oxygen and chest compressions).
4. If the person does not start breathing in 2-3 minutes, give the second dose of naloxone.  
5. Naloxone lasts for 30 to 90 minutes. If the person stops breathing again, give the second dose of naloxone. Stay with the person until help arrives. 
6. If the person is breathing but not awake, put the person on their side to prevent choking if they vomit.

Opioid overdoses are usually accidents. Please encourage your family, roommates, and friends to learn how to use naloxone. It could save your or someone else’s life. In the state of Colorado, the 911 Good Samaritan law protects citizens from arrest, charge, or prosecution when a person who is either experiencing an opioid overdose or seeing one calls 911 for help.

Classes are available to all on THIRD Thursdays of the month at 3:30-4:30PM:
A4-235 (4th floor Building A)
Rocky Mountain Regional VAMC
1700 North Wheeling St
Aurora, CO 80045

Walk-ins welcome, no appointment needed.  For questions, please call (720) 723-7418.
)
		LetterWriter(Message)
;		Msgbox Review and Sign orders. Click OK to go to next patient when ready
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
	WinActivate, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe
	GotoCPRSPage("Med")
	WinMenuSelectItem, VistA CPRS ahk_class TfrmFrame ahk_exe CPRSChart.exe, , Action, New ;Selects New Med
	WinWait, Outpatient Medications ahk_class TfrmODMeds
	
;	ControlFocus, TEdit1, Outpatient Medications ahk_class TfrmODMeds
;	ControlSetText, TEdit1, NALOXONE RESCUE SOLN
	ControlSend, TEdit1, NALOXONE R

	Loop {
	ControlGetText, MedText, TEdit1
	If InStr(MedText, "NALOXONE RESCUE SOLN,SPRAY,NASAL ")
		break
	sleep 20
	}

/*	ControlGet, Value, Enabled,, Tbutton2, Outpatient Medications ahk_class TfrmODMeds)
	while !Value{
		sleep 50
		msgbox % value
		ControlGet, Value, Enabled,, Tbutton2, Outpatient Medications ahk_class TfrmODMeds)
	}	
	*/
	; RESCUE SOLN,SPRAY,NASAL
	;send NALOXONE RESCUE SOLN,SPRAY,NASAL
	

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

	WinWait, Outpatient Medications ahk_class TfrmODMeds, Patient Instruction
	ControlSetText, TORComboEdit8, 1 SPRAY OF 4MG/SPRAY
	ControlSetText, TCaptionEdit3, 30
	ControlSetText, TCaptionEdit2, 1
	ControlSetText, TCaptionEdit1, 1
;	sleep 500
	send {down} ;Selects dosage
	control, check,,TButton2,A
;	ControlFocus, TButton2, Outpatient Medications ahk_class TfrmODMeds
;	send {enter}
	winwaitclose, Outpatient Medications ahk_class TfrmODMeds, Patient Instruction
	WinWaitActive, Outpatient Medications ahk_class TfrmODMeds, Quick Orders
	control, check,,TButton2,A
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
		WinGet, ActiveControlList, ControlList
;		msgbox % ActiveControlList
		if InStr(ActiveControlList, "TButton")
			Prefix:="T"
		else
			Prefix:=""
		
		Loop
		{
			BNumber := A_Index
			controlgettext, ButtonText, %Prefix%Button%BNumber%

			If A_Index >5
			{
;				msgbox Can't find Ok Button - %Prefix%Button%BNumber%
				BNumber := ""
				break
			}
		} until (InStr(ButtonText,"ok") or InStr(ButtonText,"yes") or InStr(ButtonText,"close") or InStr(ButtonText,"Accept") )
		if BNumber
			control, check,, %Prefix%Button%BNumber%
		WinWaitClose
		return
	}
	
	BusyPause()

return