'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: syslog
'
' AUTHOR: Nell Beatty , Computer Science
' DATE  : 4/23/2002
'
' COMMENT: This script is designed to be used in conjunction with the process
' scheduler in Windows.  It will launch this script at 11:59pm to move that
' day's application log into the stored log directory.
' Also to be used is the system log script, which will do the same thing
' but for the system log
'
'==========================================================================
Dim MyDate
MyDate = formatdatetime(Date,2)

Function ReplaceTest(initStr,patrn, replStr)
  Dim regEx, str1					' Create variables.
  str1 = initStr
  Set regEx = New RegExp				' Create regular expression.
  regEx.Pattern = patrn				' Set pattern.
  regEx.IgnoreCase = True				' Make case insensitive.
  ReplaceTest = regEx.Replace(str1, replStr)	' Make replacement.
End Function

dim slash
slash= (ReplaceTest(MyDate, "/", "_"))

slash=(ReplaceTest(slash, "/", "_"))


dim mylog
mylog = "applog"+ slash
mylog = "d:\logs\" + mylog




set owshshell=wscript.CreateObject	("wscript.shell")
owshshell.run("eventvwr.msc")
wscript.Sleep	(500)
owshshell.AppActivate ("Event Viewer")
owshshell.SendKeys	("{UP}")
owshshell.SendKeys	("{UP}")
owshshell.SendKeys	("%+a")
owshshell.SendKeys	("{DOWN}") 
owshshell.SendKeys	("{DOWN}")
owshshell.SendKeys	("{DOWN}")
owshshell.SendKeys	("{ENTER}")
owshshell.SendKeys	("{ENTER}")
owshshell.SendKeys	(mylog)
owshshell.SendKeys	("{ENTER}")
owshshell.SendKeys	("%+{F4}")