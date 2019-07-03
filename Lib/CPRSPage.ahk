CPRSPage(byref GoalTxt) ; Verifies if correct tab is active on CPRS (Wintext not accurate occasionally on Tab labels)
{
  If InStr(GoalTxt, "Med")
    GoalTxt:="Med"
  else If InStr(GoalTxt, "Lab")
    GoalTxt:="Lab" 
  else If InStr(GoalTxt, "DC")
    GoalTxt:="Discharge"   
  
  WinGetText, CPRSText, VistA CPRS in use
  WinTxtUpto2ndLine := SubStr(CPRSText, 1, InStr(CPRSText, "`n", , , 2))
  return InStr(WinTxtUpto2ndLine, GoalTxt)
}