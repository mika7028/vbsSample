Dim objRegExp       ' 正規表現オブジェクト
Dim strRepBefore    ' 置換前の文字列
Dim strRepAfter     ' 置換後の文字列

Set objRegExp = New RegExp
objRegExp.Pattern = "[A-Z]"
strRepBefore = "E-Mail: tipstest@whitire.com !!"
strRepAfter = objRegExp.Replace(strRepBefore, "*****")

WScript.Echo "置換後の文字列は " & strRepAfter & " です。"

Set objRegExp = Nothing