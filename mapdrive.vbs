Option Explicit
CreateNetwork_Conn

Sub CreateNetwork_Conn()
On Error Resume Next
Dim strDrive
Dim strShare
Dim strUserName
Dim strPassword
Dim wshNetwork

    strDrive = "Y:"
    strShare = "\\homehost\"
    strUserName = ""
    strPassword = ""

    Set wshNetwork = WScript.CreateObject("WScript.Network")

    wshNetwork.MapNetworkDrive strDrive, strShare, , strUserName, strPassword

End Sub