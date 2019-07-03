;Tries to pause script while target program is busy, thinking, etc. Also can wait for a Window to Close (IsWinClosedGlobal) or for a Window to stay active uninterrupted (IsWinOpenGlobal)
;BusyPause(IsWinClosedGlobal:="",IsWinOpenGlobal:="")
BusyPause()
{
	Start := A_TickCount
;	Stall:=(((A_Cursor = AppStarting) or (A_Cursor = Wait)) or (IsWinClosedGlobal and WinExist(IsWinClosedGlobal)) or (IsWinOpenGlobal and !WinActive(IsWinOpenGlobal)))

;	while ((((A_Cursor = AppStarting) or (A_Cursor = Wait)) or (IsWinClosedGlobal and WinExist(IsWinClosedGlobal)) or (IsWinOpenGlobal and !WinActive(IsWinOpenGlobal))) or ((A_TickCount-Start) < 100))
	while (((A_Cursor = AppStarting) or (A_Cursor = Wait)) or (IsWinClosedGlobal and WinExist(IsWinClosedGlobal)) or (IsWinOpenGlobal and !WinActive(IsWinOpenGlobal)))
	{
		If IsWinOpenGlobal
			WinActivate, %IsWinOpenGlobal%

		If IsWinClosedGlobal
			WinWaitClose, %IsWinClosedGlobal%

;		If (((A_Cursor = AppStarting) or (A_Cursor = Wait)) or (IsWinClosedGlobal and WinExist(IsWinClosedGlobal)) or (IsWinOpenGlobal and !WinActive(IsWinOpenGlobal)))
;			Start := A_TickCount

		sleep 200

	}
	IsWinClosedGlobal := ""
	IsWinOpenGlobal := ""
}