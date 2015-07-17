'===========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: Create Hogwarts User Accounts 
'
' AUTHOR: Systems Staff , Computer Sciencel
' DATE  : 6/09/2002
'
'==========================================================================

' open the computer manager to add a new user
Set owshshell=wscript.CreateObject	("wscript.shell")
owshshell.run("compmgmt.msc")
wscript.Sleep	(2500)
owshshell.SendKeys	("l")
owshshell.SendKeys	("{RIGHT}")
wscript.Sleep (500)
owshshell.SendKeys	("u")
wscript.Sleep (500)

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
' open a file that contains user information
Set oFile = oFileSystem.OpenTextFile("classroll.txt")

iA = 0
Do While iA < 2
  sLine = oFile.ReadLine
  sLinef = LTrim(sLine)
  sLinef = RTrim(sLinef)
  if Len(sLinef) = 10 And Left(sLinef,6) = "Term: " Then 
    sLine = oFile.ReadLine
    sClassNum = LTrim(sLine)
    sClassNum = RTrim(sClassNum)
    sClassNum = Right(sClassNum,Len(sClassNum)-8)
    sLine = oFile.ReadLine
    sLine = oFile.ReadLine
    sClassTitle = LTrim(sLine)
    sClassTitle = RTrim(sClassTitle)
    sClassTitle = Right(sClassTitle,Len(sClassTitle)-6)
    sClassDescription = sClassNum + sClassTitle
    sLine = oFile.ReadLine
    sLine = oFile.ReadLine
    sEnrollment = LTrim(sLine)
    sEnrollment = RTrim(sEnrollment)
    sEnrollment = Right(sEnrollment,3)
    iNo = Int(sEnrollment)
  End If
  if sLine = "=========  ====================  =======  =======  =====" Then
   iA = iA + 1
  End If
Loop

iLoopc = 0
Do While iLoopc < iNo
  sLine = oFile.ReadLine
  sLine1 = Left(sLine,40)
  sLine2 = Right(sLine1,29)
  sLine = RTrim(sLine2)
  iTemp = InStrRev(sLine,"  ")
  iTemp = Len(sLine) - iTemp
  sUserID = Right(sLine,iTemp-1)
  ' check to see if account exists
  Set oUserExists = oFileSystem.OpenTextFile("existing.txt")
  sE = oUserExists.ReadLine
  iContr = 0
  iA = 0
  Do While oUserExists.AtEndOfLine <> True
  sE = oUserExists.ReadLine
  aE = Split(sE)
  sE = aE(0)
  aE = Split(sE,"	")
  sE = aE(0)
  if StrComp(sUserID,sE,1) = 0 Then
    iContr = 1
    iA = iA - 1
    Exit Do
  End If 
  iA = iA + 1
  Loop
      
  if iContr = 0 Then
    ' there is no account anyways
  Else
    iLc = 0
    owshshell.SendKeys	("{TAB}")
    Do While iLc <= iA
      owshshell.SendKeys	("{DOWN}")
      iLc = iLc + 1
    Loop     
    wscript.Sleep (500)
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
  iLoopc = iLoopc + 1
Loop