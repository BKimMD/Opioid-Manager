CPRSPage(GoalTxt)
{
  WinGetText, CPRSText, VistA CPRS in use
  WinTxt2ndLine := SubStr(CPRSText, 1, InStr(CPRSText, "`n", , , 2))
  ;FirstNewLine := InStr(CPRSText, "`n")+1, InStr(CPRSText, "`n", ,FirstNewLine)-FirstNewLine-1)
  return InStr(WinTxt2ndLine, GoalTxt)
  ;=GoalTxt
}




/*
GetPatientSSN()
{
  ; this is probably where we want to validate CPRS instance
  PatientInfoInpatient := 0 ; clear inpatient setting
  VarSetCapacity(CPRSText, 128000)
  WinGetText, CPRSText, VistA CPRS in use by
  ;clipboard := CPRSText
  txtarray := []
  while strlen(CPRSText) > 0
  {
    txtArray[A_Index] := popline(CPRSText)
  }
  i1 := txtArray.length()
  while (strlen(trim(txtArray[i1]))=0)
  {
    i1 := i1-1
  } ; make sure any trailing blank lines are removed
  ;i=i-1 ; one more
  
  ;Let's find the line containing patient info, i1 should be last line w/pt name
  test := instr(txtarray[i1],"(OUTPATIENT)") + instr(txtarray[i1],"(INPATIENT)")
  ;msgbox, 4096, Checking Consult Conent line %i1%, % txtarray[i1] "`n" test
  while (test = 0)
  {
    i1 := i1-1
	;msgbox, 4096, Checking Consult Conent line %i1%, % txtarray[i1]
	if (i1 > 36) ; keep looking
	  test := instr(txtarray[i1],"(OUTPATIENT)") + instr(txtarray[i1],"(INPATIENT)")
	else
    {
      i1 := 0  ; failed to find
	  msgbox, 4096, Consult Read Error, Unable to read consult, make sure you have selected the correct consult.
      Break
    }	  
  }
  if (i1 > 0) ; found it, line [i1] should be last,first (INPATIENT) or (OUTPATIENT)
  { 
    ;msgbox, 4160, CPRS Patient Name Record Read, % txtarray[i1],4
	PatientInfoName := strreplace(strreplace(txtarray[i1],chr(13),""),chr(10),"")
	;msgbox, 4096, Should be ssn, % txtArray[i1-1]
    PatientInfoSSN := strreplace(txtarray[i1-1],"-","")
    PatientInfoSSN := strreplace(PatientInfoSSN,chr(10),"")
    PatientInfoSSN := trim(strreplace(PatientInfoSSN,chr(13),""))
	;msgbox, 4096, SSN, %PatientInfoSSN%
    PatientInfoDOB := substr(txtarray[i1-2],1,12)
	PatientInfoDOB := strreplace(strreplace(PatientInfoDOB,chr(13),""),chr(10),"")
	;msgbox, 4096, Patient DOB, % "DOB is: " PatientInfoDOB
    ;msgbox % txtarray[i1-4] "-4`n" txtarray[i1-3] "-3`n" txtarray[i1-2] "-2`n" txtarray[i1-1] "-1"
  
    if instr(PatientInfoName,"(OUTPATIENT)")
    {  ;this is an outpatient
      PatientInfoName := strreplace(PatientInfoName,"(OUTPATIENT)","")
	  
    }
    else if instr(PatientInfoName,"(INPATIENT)")
    { ;this is an inpatient
      PatientInfoName := strreplace(PatientInfoName,"(INPATIENT)","")
	  PatientInfoInpatient := 1
    }
    else  ; something is wrong
    {
      PatientInfoName := ""
	  PatientInfoSSN := ""
	  PatientInfoDOB := ""
    }
	;msgbox, 4096, Patient Info, % PatientInfoName "`nSSN: " PatientInfoSSN "`nDOB: " PatientInfoDOB
  }
  else
  {
      PatientInfoName := ""
	  PatientInfoSSN := ""
	  PatientInfoDOB := ""
  }
  ; this message box returns nothing
  ;msgbox 262144,, %PatientInfoName%  %PatientInfoDOB%  %PatientInfoSSN%
  Return
}
*/