//
// IELauncher.cs
// Windows11でInternetExplorerを起動する。
//
// コマンドラインにURLやファイルを指定すると開く。
//

using Microsoft.VisualBasic;

var ie = (SHDocVw.InternetExplorer)Interaction.CreateObject("InternetExplorer.Application");

string[] cmds = Environment.GetCommandLineArgs();

if(cmds.Length > 1) ie.Navigate(cmds[1]);

ie.Visible = true;

// ◆Createしただけだと最前面に来ないようだ。
// ◆IEは複数開いてもプロセス数は常に一定のようで、0番目を最前面にすれば良いのかもしれないが、一応全部アクティベートする。
// ◆うまくいったりいかなかったり。
System.Diagnostics.Process.GetProcessesByName("iexplore")
	.ToList()
	.ForEach(p => Interaction.AppActivate(p.Id));
