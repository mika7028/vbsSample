Dim objRegExp       ' ���K�\���I�u�W�F�N�g
Dim strRepBefore    ' �u���O�̕�����
Dim strRepAfter     ' �u����̕�����

Set objRegExp = New RegExp
objRegExp.Pattern = "[A-Z]"
strRepBefore = "E-Mail: tipstest@whitire.com !!"
strRepAfter = objRegExp.Replace(strRepBefore, "*****")

WScript.Echo "�u����̕������ " & strRepAfter & " �ł��B"

Set objRegExp = Nothing