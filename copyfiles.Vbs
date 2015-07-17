' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM) 
' 
' NAME: copy firewall
' 
' AUTHOR: Nell Beatty , Computer Science 
' DATE  : 11/07/2002
' 
' COMMENT: This script is designed to be used in conjunction with the process 
' scheduler in Windows.  It will launch this script at 11:59pm to move that 
' day's firewall log into the stored log directory.
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

Dim mylog
mylog = "firewall" + slash


Function DistributeFiles(strSourceDir, strDestDir)

     Dim lfsObject
     Dim lFileList
     Dim lFolder
     Dim lFile

     On Error Resume Next
     Set lfsObject = CreateObject("Scripting.FileSystemObject")
     Set lFolder = lfsObject.GetFolder(strSourceDir)
     Set lFileList = lFolder.Files
     For Each lFile In lFileList
          Dim tempname
          tempname = lFile.Name + mylog
          lfsObject.CopyFile strSourceDir & "\" & lFile.Name, strDestDir & "\" & tempname           
     Next     

End Function

Dim mydir
Dim logdir

mydir = "c:\Winnt\Internet Logs"
logdir = "d:\logs\firewall"

DistributeFiles mydir, logdir


