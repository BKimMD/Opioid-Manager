;Confirms Tab Title sent = Tab on CPRS
GotoCPRSPage(TabTitle)
{
;	Count:=0
	while !CPRSPage(TabTitle)
	{

;		Count:=Count + 1

		if InStr(TabTitle, "Order")
			send ^o
		else if InStr(TabTitle, "Report")
			send ^r
		else if InStr(TabTitle, "Lab")
			send ^l
		else if InStr(TabTitle, "Note")
			send ^n
		else if InStr(TabTitle, "Med")
			send ^m
		else if InStr(TabTitle, "Consult")
			send ^t
		else if InStr(TabTitle, "Problem")
			send ^p
		else if InStr(TabTitle, "Cover")
			send ^s
		else if InStr(TabTitle, "Discharge")
			send ^d
		else if InStr(TabTitle, "Surg")
			send ^u
		else
		{
			msgbox Unable to find -%TabTitle%- CPRS Tab
			break
		}
		
		BusyPause()
		/*
		if Count > 50
			{
			MsgBox GotoCPRS Tab finder Timed Out
			break
			}
			*/
	}
}