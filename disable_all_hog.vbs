'===========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: Disable Hogwarts User Accounts 
'
' AUTHOR: Systems Staff , Computer Sciencel
' DATE  : 6/09/2002
'
'==========================================================================

' open the computer manager
Set owshshell=wscript.CreateObject	("wscript.shell")
owshshell.run("compmgmt.msc")
wscript.Sleep	(2500)
owshshell.SendKeys	("l")
owshshell.SendKeys	("{RIGHT}")
wscript.Sleep (500)
owshshell.SendKeys	("u")
wscript.Sleep (500)

' export the list of existing users to a file called existing.txt 
' export the list of existing users to a file called existing.txt 
owshshell.SendKeys	("%")
wscript.Sleep (500)
owshshell.SendKeys	("~")
wscript.Sleep (500)
owshshell.SendKeys	("l")
wscript.Sleep (500)
owshshell.SendKeys	("c:\admspt\existing.txt")
owshshell.SendKeys	("~")
wscript.Sleep (500)
owshshell.SendKeys	("{TAB}")
wscript.Sleep (500)
owshshell.SendKeys	("~")
wscript.Sleep (500)	

Set oFileSystem = WScript.CreateObject("Scripting.FileSystemObject")

  Set oUserExists = oFileSystem.OpenTextFile("existing.txt")
  sE = oUserExists.ReadLine
  iA = 0
  Do While oUserExists.AtEndOfLine <> True
  sE = oUserExists.ReadLine
  aE = Split(sE)
  sE = aE(0)
  aE = Split(sE,"	")
  sE = aE(0)
  if StrComp("administrator",sE,1) = 0 Then
  '  do nothing to administrator account
  Else
    iLc = 0
    owshshell.SendKeys	("{TAB}")
    Do While iLc < iA
      owshshell.SendKeys	("{DOWN}")
      iLc = iLc + 1
    Loop  
    ' disable the account  
    owshshell.SendKeys	("%")
    owshshell.SendKeys	("~")
    owshshell.SendKeys	("r")
    wscript.Sleep (500)
    owshshell.SendKeys	("{TAB}")
    owshshell.SendKeys	("{TAB}")
    owshshell.SendKeys	("b")
    owshshell.SendKeys	("=")
    owshshell.SendKeys	("~")
    wscript.Sleep (500) 
    owshshell.SendKeys	("{TAB}")
    owshshell.SendKeys	("{TAB}")
    wscript.Sleep (500)
  End If
  iA = iA + 1
Loop