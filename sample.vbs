'�錾
Dim inputYear
Dim inputMonth
Dim inputDay
Dim fullDate
Dim outputYear
Dim outputJapaneseCalendar

'�ϐ��i�[
inputYear = 1866
inputMonth = 5
inputDay = 1

if inputYear < 1868 then
	msgbox "�G���[�ł��B"
else

'���Ɠ��t���ꕶ���̏ꍇ�́u0�v�ۊ�
if len(inputMonth) = 1 then
	inputMonth = "0" & inputMonth
End If
if len(inputDay) = 1 then
	inputDay = "0" & inputDay
End If

'����ƌ��Ɠ��t���q����
fullDate = inputYear & inputMonth & inputDay

'��������
if CCur(fullDate) <= 19120929 then
	outputJapaneseCalendar = "����"
	outputYear = inputYear - 1868 + 1
	if outputYear = 1 then
		outputYear = "��"
	end if 
end if

'�吳����
if CCur(fullDate) > 19120929 and CCur(fullDate) <= 19261224 then
	outputJapaneseCalendar = "�吳"
	outputYear = inputYear - 1912 + 1
	if outputYear = 1 then
		outputYear = "��"
	end if 
end if 

'���a����
if CCur(fullDate) > 19261224 and CCur(fullDate) <= 19890107 then
	outputJapaneseCalendar = "���a"
	outputYear = inputYear - 1926 + 1
	if outputYear = 1 then
		outputYear = "��"
	end if 
end if 

'��������
if CCur(fullDate) > 19890107 and CCur(fullDate) <= 20190430 then
	outputJapaneseCalendar = "����"
	outputYear = inputYear - 1989 + 1
	if outputYear = 1 then
		outputYear = "��"
	end if 
end if 

'�ߘa����
if CCur(fullDate) > 20190430 then
	outputJapaneseCalendar = "�ߘa"
	outputYear = inputYear - 2019 + 1
	if outputYear = 1 then
		outputYear = "��"
	end if 
end if 
 
fullDate = outputJapaneseCalendar & outputYear & "�N" & inputMonth & "��"  & inputDay & "��" 
msgbox fullDate

end if
'��ƂȂ���t
'inputDate = !���!
'���Z�������N��
'num = !�N����!

'���t�����Z����
'outputDate = DateAdd("yyyy", num, inputDate)

'UMS�̕ϐ��Ɋi�[
'SetUMSVariable $���ʊi�[��$, outputDate