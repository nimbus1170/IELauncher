Rem 元プログラム
Rem Windows11でInternet Exploer11を起動する方法 - Qiita
Rem https://qiita.com/lumin/items/bd2a8e90efb4010e2804

Set objArgs = WScript.Arguments

Set objIE = CreateObject("InternetExplorer.Application")

If objArgs.Count = 0 Then
    objIE.Navigate "https://www.msn.com/ja-jp/"
Else
    objIE.Navigate objArgs(0)
End If

objIE.Visible = true

rem objIE.AppActivate
