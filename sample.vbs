'éŒ¾
Dim inputYear
Dim inputMonth
Dim inputDay
Dim fullDate
Dim outputYear
Dim outputJapaneseCalendar

'•Ï”Ši”[
inputYear = 1866
inputMonth = 5
inputDay = 1

if inputYear < 1868 then
	msgbox "ƒGƒ‰[‚Å‚·B"
else

'Œ‚Æ“ú•t‚ªˆê•¶š‚Ìê‡‚Íu0v•ÛŠÇ
if len(inputMonth) = 1 then
	inputMonth = "0" & inputMonth
End If
if len(inputDay) = 1 then
	inputDay = "0" & inputDay
End If

'¼—ï‚ÆŒ‚Æ“ú•t‚ğŒq‚°‚é
fullDate = inputYear & inputMonth & inputDay

'–¾¡ˆ—
if CCur(fullDate) <= 19120929 then
	outputJapaneseCalendar = "–¾¡"
	outputYear = inputYear - 1868 + 1
	if outputYear = 1 then
		outputYear = "Œ³"
	end if 
end if

'‘å³ˆ—
if CCur(fullDate) > 19120929 and CCur(fullDate) <= 19261224 then
	outputJapaneseCalendar = "‘å³"
	outputYear = inputYear - 1912 + 1
	if outputYear = 1 then
		outputYear = "Œ³"
	end if 
end if 

'º˜aˆ—
if CCur(fullDate) > 19261224 and CCur(fullDate) <= 19890107 then
	outputJapaneseCalendar = "º˜a"
	outputYear = inputYear - 1926 + 1
	if outputYear = 1 then
		outputYear = "Œ³"
	end if 
end if 

'•½¬ˆ—
if CCur(fullDate) > 19890107 and CCur(fullDate) <= 20190430 then
	outputJapaneseCalendar = "•½¬"
	outputYear = inputYear - 1989 + 1
	if outputYear = 1 then
		outputYear = "Œ³"
	end if 
end if 

'—ß˜aˆ—
if CCur(fullDate) > 20190430 then
	outputJapaneseCalendar = "—ß˜a"
	outputYear = inputYear - 2019 + 1
	if outputYear = 1 then
		outputYear = "Œ³"
	end if 
end if 
 
fullDate = outputJapaneseCalendar & outputYear & "”N" & inputMonth & "Œ"  & inputDay & "“ú" 
msgbox fullDate

end if
'Šî€‚Æ‚È‚é“ú•t
'inputDate = !Šî€“ú!
'‰ÁZ‚µ‚½‚¢”N”
'num = !”N”·!

'“ú•t‚ğ‰ÁZ‚·‚é
'outputDate = DateAdd("yyyy", num, inputDate)

'UMS‚Ì•Ï”‚ÉŠi”[
'SetUMSVariable $Œ‹‰ÊŠi”[æ$, outputDate