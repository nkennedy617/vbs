'==========================================================================
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
' AUTHOR: Systems Staff , Computer Science
' COMMENT: open and close Retrospect application
' DATE  : 5/29/2002
'==========================================================================

set owshshell=wscript.CreateObject	("wscript.shell")
owshshell.run("Retrospect.exe")
wscript.Sleep	(15000)
owshshell.AppActivate ("Retrospect")
wscript.Sleep	(500)
owshshell.SendKeys  ("%")
owshshell.SendKeys	("~")
owshshell.SendKeys	("x")
wscript.Sleep	(500)
owshshell.SendKeys	("~")