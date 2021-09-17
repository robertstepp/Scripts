' ------------------------------------------------------------------------
' Scriptname: Reset-LAPEnc.vbs
' Author: Isaac Brumley
' Initial Release 3/25/2010
' Updated: 8/11/2011

' ------------------------------------------------------------------------
' IMPLEMENTATION NOTES:
'
' Encode the script before publishing to group policy and lock down the 
' permissions to "Domain Computers" only.  You can add Administrators 
' when you need to edit the script in the GPO.

' ------------------------------------------------------------------------
OPTION EXPLICIT		' No undefined variables allowed!

' ------------------------------------------------------------------------
On Error Resume Next	' Show no errors

' ------------------------------------------------------------------------
Dim sUserSID, oWshNetwork, oUserAccount, strNewPassword, intType
Dim WshShell, sComputer, sAdminName, oUser, rc, oUserAccounts
Dim strMessage, namedArgs


'==============================================================================
' Arguments check
Set namedArgs = wscript.arguments.named
' -----------------------------------------------------------------------------
if namedArgs.exists("Encrypt") Then
	strNewPassword = namedArgs.item("Encrypt")
	strNewPassword = Encrypt(strNewPassword) 
	strNewPassword = Encrypt(strNewPassword) 
	WScript.Echo strNewPassword 
	WScript.Quit
end If

' -----------------------------------------------------------------------------
if namedArgs.exists("Decrypt") Then
	strNewPassword = namedArgs.item("Decrypt")
	strNewPassword = Decrypt(strNewPassword) 
	strNewPassword = Decrypt(strNewPassword) 
	WScript.Echo strNewPassword 
	WScript.Quit
end If

' -----------------------------------------------------------------------------
If WScript.Arguments.Unnamed.Count = 1 Then
	strNewPassword = WScript.Arguments.Unnamed(0)
	strNewPassword = Decrypt(strNewPassword) ' Using double encrypted PW
	strNewPassword = Decrypt(strNewPassword) ' Using double encrypted PW
End If

If WScript.Arguments.Count = 0 Then
	Wscript.Echo "No arguments passed"
	WScript.quit
End If
' -----------------------------------------------------------------------------
rc = 0		' Initialize the return code variable

' ------------------------------------------------------------------------
' Create required objects here.
Set WshShell = CreateObject("WScript.Shell")
Set oWshNetwork = CreateObject("WScript.Network")

' ------------------------------------------------------------------------
sComputer = oWshNetwork.ComputerName	' Get the local computer name
sAdminName = GetAdministratorName()	' Get the true local administrators name


' ------------------------------------------------------------------------
Set oUser = GetObject("WinNT://" & sComputer & "/" & sAdminName & ",user")
	oUser.SetPassword strNewPassword
	rc = oUser.SetInfo
	If rc = 0 Then
		strMessage = "Password successfully changed for " & sAdminName
		Logevent strMessage,8
	Else
		strMessage = "Password reset failed for " & sAdminName
		Logevent strMessage,16
	End If
	'On Error Goto 0

' ------------------------------------------------------------------------
' Cleanup objects used
Set WshShell = nothing
Set oWshNetwork = nothing
Set oUser = nothing

WScript.quit(0)		' Logical end of script


' ------------------------------------------------------------------------
' This function will obtain the current Administrator account name no matter
' what it was renamed to.
' ------------------------------------------------------------------------
Function GetAdministratorName()
Set oUserAccounts = GetObject("winmgmts://" & oWshNetwork.ComputerName & _
	"/root/cimv2").ExecQuery("Select Name, SID from Win32_UserAccount" _
	& " WHERE Domain = '" & oWshNetwork.ComputerName & "'")
'On Error Resume Next
For Each oUserAccount In oUserAccounts
	If Left(oUserAccount.SID, 9) = "S-1-5-21-" And Right(oUserAccount.SID, 4) = "-500" Then
		GetAdministratorName = oUserAccount.Name
	Exit For
	End if
Next

' Cleanup objects used
Set oUserAccounts = nothing

End Function

' ------------------------------------------------------------------------
' Log event to Application Event Log
' ------------------------------------------------------------------------
Function Logevent(strMessage,intType)
	WshShell.LogEvent intType, strMessage
End Function
' ------------------------------------------------------------------------
' *****************************************************************************
Private Function Encrypt(ByVal string)
' Provided by: http://www.psacake.com/web/func/encrypt_function.htm
' *****************************************************************************
	Dim x, i, tmp
	For i = 1 To Len( String )
		x = Mid( string, i, 1 )
		tmp = tmp & Chr( Asc( x ) + 1 )
	Next
	tmp = StrReverse( tmp )
	Encrypt = tmp
End Function

' *****************************************************************************
Private Function Decrypt(ByVal encryptedstring)
' Provided by: http://www.psacake.com/web/func/decrypt_function.htm
' *****************************************************************************
	Dim x, i, tmp
	encryptedstring = StrReverse( encryptedstring )
	For i = 1 To Len( encryptedstring )
		x = Mid( encryptedstring, i, 1 )
		tmp = tmp & Chr( Asc( x ) - 1 )
	Next
	Decrypt = tmp
End Function
' *****************************************************************************
