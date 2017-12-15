StringSplit, FldrArray, A_WorkingDir, \

Railroad = %FldrArray2%
YearFolder = %FldrArray3%
XrlNum = %FldrArray4%
LocationFolder = %FldrArray5%

If Railroad = CSX
	Run, "P:\Designer\Aaron\docCreator\Doc Creator.hta" "1"
else If Railroad = NS
	Run, "P:\Designer\Aaron\docCreator\Doc Creator.hta" "1"
else	
	return 

WinWaitActive Which Railroad?, ,3

If Railroad = CSX
	Send {Space}
else
	Send {Right}{Space}
	
WinWaitActive Job Year?, ,3
	Send %YearFolder%{Enter}
	
WinWaitActive Xorail Number?, ,3
	Send %XrlNum%{Enter}
	
WinWaitActive Location Folder Name?, ,3
	Send %LocationFolder%{Enter}

Loop, * info worksheet.xlsm
	Filename = %A_LoopFileName%

StringReplace, LocationName, Filename, info worksheet.xlsm,

WinWaitActive Location Name?, ,3
	Send % LTrim(LocationName,"")
	Send {Enter}


