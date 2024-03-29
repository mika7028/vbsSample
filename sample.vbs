'宣言
Dim inputYear
Dim inputMonth
Dim inputDay
Dim fullDate
Dim outputYear
Dim outputJapaneseCalendar

'変数格納
inputYear = 1866
inputMonth = 5
inputDay = 1

if inputYear < 1868 then
	msgbox "エラーです。"
else

'月と日付が一文字の場合は「0」保管
if len(inputMonth) = 1 then
	inputMonth = "0" & inputMonth
End If
if len(inputDay) = 1 then
	inputDay = "0" & inputDay
End If

'西暦と月と日付を繋げる
fullDate = inputYear & inputMonth & inputDay

'明治処理
if CCur(fullDate) <= 19120929 then
	outputJapaneseCalendar = "明治"
	outputYear = inputYear - 1868 + 1
	if outputYear = 1 then
		outputYear = "元"
	end if 
end if

'大正処理
if CCur(fullDate) > 19120929 and CCur(fullDate) <= 19261224 then
	outputJapaneseCalendar = "大正"
	outputYear = inputYear - 1912 + 1
	if outputYear = 1 then
		outputYear = "元"
	end if 
end if 

'昭和処理
if CCur(fullDate) > 19261224 and CCur(fullDate) <= 19890107 then
	outputJapaneseCalendar = "昭和"
	outputYear = inputYear - 1926 + 1
	if outputYear = 1 then
		outputYear = "元"
	end if 
end if 

'平成処理
if CCur(fullDate) > 19890107 and CCur(fullDate) <= 20190430 then
	outputJapaneseCalendar = "平成"
	outputYear = inputYear - 1989 + 1
	if outputYear = 1 then
		outputYear = "元"
	end if 
end if 

'令和処理
if CCur(fullDate) > 20190430 then
	outputJapaneseCalendar = "令和"
	outputYear = inputYear - 2019 + 1
	if outputYear = 1 then
		outputYear = "元"
	end if 
end if 
 
fullDate = outputJapaneseCalendar & outputYear & "年" & inputMonth & "月"  & inputDay & "日" 
msgbox fullDate

end if
'基準となる日付
'inputDate = !基準日!
'加算したい年数
'num = !年数差!

'日付を加算する
'outputDate = DateAdd("yyyy", num, inputDate)

'UMSの変数に格納
'SetUMSVariable $結果格納先$, outputDate