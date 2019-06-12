CSVtoArray(file)
{
	NamesArray := []
	arrayOfHeaders := []
	arrayOfData := []
	If !FileExist(file)
	{
		Msgbox Batch .CSV file not found, exiting
		ExitApp
	}
	else
	{
		Loop, Read, %file%
		{
				NamesArray[A_Index] := {}	;due to the way AHK works, you need to separate object references. http://www.autohotkey.com/board/topic/77221-associative-array-of-objects-help/
				NamesArray[A_Index]  := StrSplit(A_LoopReadLine, ",", """") ;Parsing each line of CSV
				If NamesArray[A_Index, 1] = ""
				{
					NamesArray.Delete(A_Index)
					break
				}
				Pack := "000000000"
				NamesArray[A_Index, 3] := SubStr(Pack, 1, StrLen(Pack) - StrLen(NamesArray[A_Index, 3])) . NamesArray[A_Index, 3]
				NamesArray[A_Index, 13] := DateParse(NamesArray[A_Index, 13])"000000"
				NamesArray[A_Index, 19] := DateParse(NamesArray[A_Index, 19])"000000"
				
	;			NamesArray[A_Index, 3] := SubStr(NamesArray[A_Index, 3], -3, 4)

	;			msgbox % NamesArray[A_Index, 3]
	;			msgbox % array[2][1]
				;make a dictionary of {HeaderName: DataItem}
	;			subdict := {}	;This line is needed create a new object reference to subdict in memory
	;			For index, value in arrayOfData	;for all columns in one row, create dict of {HeaderName: DataItem, HeaderName: DataItem}
	;				subdict[index] := value
	;			array[A_Index] := subdict			;==array.Insert(A_Index - 1, subdict)	;subdict := {} is needed above else this line just adds the reference to the subdict, not the new values
			
		}
	return NamesArray
	}
}