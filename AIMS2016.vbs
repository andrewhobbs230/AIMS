'////////////////////////////////////////////////
'/  AIMS2016.vbs                                   
'/   (c)2016 Bullitt County Schools             
'/    A.Hobbs                                   
'/     v1.0                                     
'
'////////////////////////////////////////////////

'##################################
'#      Global Vars				  #
'##################################
cnstring = "Provider=SQLOLEDB.1;Password=******;Persist Security Info=True;User ID=******;Initial Catalog=Aimsv2;Data Source=******"
HomeDirectoryServer = "*******"
HomeDrive = "h"
PasswordLength = "4"
PreWindowsDomain = "*******"
strDomain = "*****.ketsds.net"

'##################################
'#      Main Script Calls		  #
'##################################
LogEvent "63", "New User Creation Has Began."
MainCreateNew()
LogEvent "63", "New User Creation Has Stopped."

LogEvent "63", "User Moves Have Began."
MainMoveAccounts()
LogEvent "63", "User Moves Have Stopped."

LogEvent "63", "User Quarantine Has Began."
DisableInactiveUnEnrolledStudents()
LogEvent "63", "User Quarantine Has Stopped."

LogEvent "63", "Enabling Accounts Has Began."
EnableActiveEnrolledStudents()
LogEvent "63", "Enabling Accounts Has Stopped."

'##################################
'#      Functions and Subs		  #
'##################################

Function CreateUser(fname,lname,snum,grade,grad_year,sn)
username = CreateUsername(fname,lname)
Password = CreatePassword(snum)
SiteName = GetSiteName(sn)

Set objOU = GetObject("LDAP://OU=students,dc=bullitt,dc=ketsds,dc=net")
Set objUser = objOU.Create("user", "cn="&username)
objUser.Put "sAMAccountName", username
objUser.Put "description", grad_year
objUser.Put "Physicaldeliveryofficename", SiteName
objUser.Put "givenName", fname
objUser.Put "SN", lname
objUser.Put "userPrincipalName", username&"@bullitt.ketsds.net"
objUser.Put "mail", username&"@stu.bullitt.kyschools.us"
objUser.Put "homedrive", HomeDrive
objUser.Put "homedirectory", HomeDirectoryServer & username
objUser.Put "DisplayName", lname&", "&fname
objUser.Put "Department", snum
objUser.Put "KetsCustom1", snum
objUser.SetInfo

'Return the New Username you just created
CreateUser = username
End Function

Function CreateFolder(username)
On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder(HomeDirectoryServer & username)
End Function

Function DoesUserExist(strUser)
strDomain = "bullitt.ketsds.net"

Const ADS_SCOPE_SUBTREE = 2
Set cnn = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")
cnn.Provider = "ADsDSOObject"
cnn.Open "Active Directory Provider"
Set cmd.ActiveConnection = cnn
cmd.Properties("Page Size") = 1000
cmd.Properties("Timeout") = 30
cmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
cmd.Properties("Cache Results") = False

cmd.CommandText = "SELECT name from 'LDAP://" & strDomain & "' WHERE objectCategory = 'user' AND SAMAccountName = '" & strUser & "'"

Set rs = cmd.Execute
If rs.EOF Then
    DoesUserExist = false
  Else
    DoesUserExist = true
End If

End Function

Function AddUserToAccountTable(AuthorityID, fname, lname, Location, grade, GraduationYear, CreateDate, ADUsername, OriginalPassword)
cmdtext = "INSERT INTO _MasterAccounts (AuthorityID, fname, lname, Location, grade, GraduationYear, CreateDate, ADUsername, OriginalPassword)"
cmdtext = cmdtext & " VALUES ('"&AuthorityID&"', '"&fname&"', '"&lname&"', '"&Location&"', '"&grade&"', '"&GraduationYear&"', '"&CreateDate&"', '"&ADUsername&"', '"&OriginalPassword&"')"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)
cn.Close
End Function

Function AddUserToGroup(username,number)
On Error Resume Next
GroupPath = "LDAP://cn="&GetGroupName(number)&",ou=_Groups,ou=Students,dc=bullitt,dc=ketsds,dc=net"
userPath = "LDAP://cn="&username&",ou=students,dc=bullitt,dc=ketsds,dc=net"
Set objGroup = getobject(GroupPath)
'	for each member in objGroup.members
'		if lcase(member.adspath) = lcase(userPath) then
'			exit Function
'		end if
'	next
	objGroup.Add(userPath)
End Function

Function AddUserToGroupByADGroupName(username,group)
GroupPath = "LDAP://cn="&group&",ou=_Groups,ou=Students,dc=bullitt,dc=ketsds,dc=net"
userPath = "LDAP://cn="&username&",ou=students,dc=bullitt,dc=ketsds,dc=net"
Set objGroup = getobject(GroupPath)
	for each member in objGroup.members
		if lcase(member.adspath) = lcase(userPath) then
			exit Function
		end if
	next
	objGroup.Add(userPath)
End Function

Function RemoveUserFromGroup(username,group)
GroupPath = "LDAP://cn="&group&",ou=_Groups,ou=Students,dc=bullitt,dc=ketsds,dc=net"
userPath = "LDAP://cn="&username&",ou=students,dc=bullitt,dc=ketsds,dc=net"
Set objGroup = getobject(groupPath)
	for each member in objGroup.members
		if lcase(member.adspath) = lcase(userPath) then
			objGroup.Remove(userPath)
			exit Function
		end if
	next
End Function


Function CreateUsername(fname,lname)
baseusername = fname &"."&lname
	If Len(baseusername) > 18 Then
		baseusername = Left(baseusername,17)
	End If
baseusername = Replace(baseusername, "'", "")
baseusername = Replace(baseusername, "-", "")
baseusername = Replace(baseusername, " ", "")
baseusername = Replace(baseusername, ",", "")
username = baseusername
	i=1
	Do Until UserExists = "False"
		If DoesUserExist(username) = "True" Then
			username = baseusername & i
			UserExists = "True"
		Else 
			FinalUsername = username
			UserExists = "False"
		End If
	i=i+1
	Loop
	CreateUsername = LCase(FinalUsername)
End Function


Function FixString(text)
If Len(text) > 0 Then
	FixString = Replace(text, "'", "")
Else
	'Do Something
End If
End Function

Function GetDNBySAM(sAMAccountName)
Const ADS_SCOPE_SUBTREE = 2
Set objRootDSE = GetObject("LDAP://RootDSE")
strDomain = objRootDSE.Get("DefaultNamingContext")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand = CreateObject("ADODB.Command")
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.CommandText = "SELECT distinguishedName FROM 'LDAP://" & strDomain & "' WHERE objectCategory='user' AND samAccountName = '" & sAMAccountName & "'"
Set objRecordSet = objCommand.Execute
If Not objRecordSet.EOF Then
 GetDNBySAM = objRecordSet.Fields("distinguishedName").Value
End If
End Function


Function SetHomeDirPermissions(NewUser)
Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
UserFolderPath = HomeDirectoryServer & NewUser

Set userfolder = fso.GetFolder(UserFolderPath)
	
	strCommand1 = "echo y| cacls " & UserFolderPath
	strCommand2 = " /C /G Administrators:F ""bullitt\Dist Support Admins"":F ""bullitt\Dist Staff"":F ""System"":F ""bullitt\Dist Leadership"":F " & PreWindowsDomain & "\" & NewUser & ":F"
	oShell.Run "cmd /C " & strCommand1 & " /T" & strCommand2, 0, True

End Function

Function DisableInactiveUnEnrolledStudents()
On Error Resume Next
cmdtext = "Select * from v_InactiveStudents order by adusername"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)
Do Until rs.EOF
strUsername = rs.Fields.Item("ADUsername").value
strSchoolName = rs.Fields.Item("SchoolLongName").Value
strGradYear = rs.Fields.Item("GraduationYear").Value
'strOU = rs.Fields.Item("Stu_OU").Value
If DoesUserExist(strUsername) = True Then
'LogEvent "User_Quarantined", strUsername &" - has been quarantined."
DN = GetDNBySAM(strUsername)
	Set objOU = GetObject("LDAP://" & DN) 
		objOU.accountDisabled = True
		objOU.description = "Quarantined - " & strSchoolName & " " & strGradYear
		objOU.SetInfo 
End If
	rs.MoveNext
Loop	
rs.Close
cn.Close
End Function

Function EnableActiveEnrolledStudents()
On Error Resume Next
cmdtext = "Select * from v_ActiveStudents order by adusername"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)
Do Until rs.EOF
strUsername = rs.Fields.Item("ADUsername").value
strSchoolName = rs.Fields.Item("SchoolLongName").Value
strGradYear = rs.Fields.Item("GraduationYear").Value
'strOU = rs.Fields.Item("Stu_OU").Value
DN = GetDNBySAM(strUsername)

	Set objOU = GetObject("LDAP://" & DN) 
		objOU.accountDisabled = False
		objOU.description = strGradYear
		objOU.Physicaldeliveryofficename = strSchoolName
		objOU.SetInfo 

	rs.MoveNext
Loop	
rs.Close
cn.Close
End Function


Function CreatePassword(snum)
FinalPassword = Right(snum, PasswordLength)
CreatePassword = FinalPassword
End Function

Function SetPassword(username,snum)
Password = CreatePassword(snum)
Set objUser = GetObject _
 ("LDAP://cn="&username&",ou=students,dc=bullitt,dc=ketsds,dc=net")
objUser.SetPassword Password
SetPassword = Password
End Function

Function GetSiteName(sn)
cmdtext = "Select * from Locations where SchoolNumber = '"&sn&"'"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)

Do Until rs.EOF
	GetSiteName = rs.Fields.Item("SchoolLongName").Value
rs.MoveNext
Loop	
rs.Close
cn.Close

End Function

Function GetGroupName(number)
cmdtext = "Select * from Locations where SchoolNumber = '"&number&"'"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)

Do Until rs.EOF
	GetGroupName = rs.Fields.Item("Stu_ADgroup1").Value
rs.MoveNext
Loop	
rs.Close
cn.Close

End Function

Function UpdateSNInAIMS(ad_username, NewSN)
cmdtext = "Update _MasterAccounts Set Location ='"&NewSN&"' where ADUsername ='"&ad_username&"'"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)
End Function


Function LogEvent(LogType, LogText)
strLogType = FixString(LogType)
strLogText = FixString(LogText)
cmdtext = "INSERT INTO Log (LogType, LogText, LogOwner)"
cmdtext = cmdtext & " VALUES ('"&strLogType&"', '"&strLogText&"', '999999')"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)
cn.Close
End Function


Function SendNewUserEmail(Username,sn)
strSubject = "AIMS - "&Username&" Created"
strMailFrom = "AIMS@bullitt.ketsds.net"
strMailTo = "*****@123.com"
strBody = "New User "&UserName&" Has Been Created"

Set myMail=CreateObject("CDO.Message")
myMail.Subject= strSubject
myMail.From= strMailFrom
myMail.To= strMailTo
mymail.TextBody = strBody
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
'Name or IP of remote SMTP server
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
="ketsmail.us"
'Server port
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") _
=25 
myMail.Configuration.Fields.Update
myMail.Send
Set myMail=Nothing
End Function

Function SendMoveUserEmail(Username)
strSubject = "AIMS - "&Username&" Created"
strMailFrom = "AIMS@bullitt.ketsds.net"
strMailTo = GetNotifyEmailAddress(sn)
strMailCC = "******@bullitt.kyschools.us"
strBody = "New User "&UserName&" Has Been Created"

Set myMail=CreateObject("CDO.Message")
myMail.Subject= strSubject
myMail.From= strMailFrom
myMail.To= strMailTo
myMail.CC = strMailCC
mymail.TextBody = strBody
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
'Name or IP of remote SMTP server
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
="ketsmail.us"
'Server port
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") _
=25 
myMail.Configuration.Fields.Update
myMail.Send
Set myMail=Nothing
End Function

Function EnableAccount(DN)
On Error Resume Next
	Set objOU = GetObject("LDAP://"& DN) 
		objOU.accountDisabled = False
		objOU.SetInfo 
End Function

Function RemoveRecordFromMoveTable(id)
cmdtext = "Delete from Move_Accounts_tmp where ID = '"&id&"'"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs1 = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring 
cn.Open 
cmd.activeconnection = cn 
rs1.cursortype = 3 
rs1.locktype = 3 
cmd.commandtext = cmdtext 
rs1.Open(cmd)
End Function

Function SendErrorEmail(ErrorText, EmailTo)
strSubject = "AIMS Error Has Been Thrown"
strMailFrom = "AIMS@bullitt.ketsds.net"
strMailTo = EmailTo
strBody = ErrorText

Set myMail=CreateObject("CDO.Message")
myMail.Subject= strSubject
myMail.From= strMailFrom
myMail.To= strMailTo
mymail.TextBody = strBody
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
'Name or IP of remote SMTP server
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
="ketsmail.us"
'Server port
myMail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") _
=25 
myMail.Configuration.Fields.Update
myMail.Send
Set myMail=Nothing
End Function

'##############################
'#   Core AIMS Function		  #
'##############################

Function MainCreateNew()
On Error Resume Next
cmdtext = "Select * from v_GetNewUsersToCreate"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring
cn.ConnectionTimeout = 0 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)

Do Until rs.EOF
	fname = FixString(rs.Fields.Item("FirstName").Value)
	mname = FixString(rs.Fields.Item("MiddleName").Value)
	lname = FixString(rs.Fields.Item("LastName").Value)
	StateID = rs.Fields.Item("StateID").Value
	grade = rs.Fields.Item("grade").Value
	number = rs.Fields.Item("number").Value
	gradyear = rs.Fields.Item("gradyear").Value
	cre_date = rs.Fields.Item("DateCreated").Value
	

NewUser = CreateUser(fname,lname,StateID,grade,gradyear,number)
Password = SetPassword(NewUser,StateID)
AddUserToGroup NewUser,number
CreateFolder NewUser
SetHomeDirPermissions NewUser
AddUserToAccountTable StateID,fname,lname,number,grade,gradyear,cre_date,NewUser,Password
SendNewUserEmail NewUser,number
LogEvent "64", NewUser&" has been created for "&number&"."



rs.MoveNext
Loop	
rs.Close
cn.Close
End Function


Function MainMoveAccounts()
On Error Resume Next
cmdtext = "Select * from v_GetMoveAccounts"
Set cn = CreateObject("adodb.connection") 
Set cmd = CreateObject("adodb.command") 
Set rs = CreateObject("adodb.recordset") 
cn.connectionstring = cnstring
cn.ConnectionTimeout = 0 
cn.Open 
cmd.activeconnection = cn 
rs.cursortype = 3 
rs.locktype = 3 
cmd.commandtext = cmdtext 
rs.Open(cmd)

Do Until rs.EOF
'Assign Fields to Vars
	ad_username = rs.Fields.Item("ad_username").Value
	AddToADGroup = rs.Fields.Item("AddToADGroup").Value
	RemoveFromADGroup = rs.Fields.Item("RemoveFromADGroup").Value
	NewPhyOffice = rs.Fields.Item("NewPhyOffice").Value
	NewSN = rs.Fields.Item("NewSN").Value
'Do the Move!
'MsgBox(ad_username)
If DoesUserExist(ad_username) = True Then
		RemoveUserFromGroup ad_username, RemoveFromADGroup
		AddUserToGroupByADGroupName ad_username, AddToADGroup
		UpdateSNInAIMS ad_username, NewSN
		Set objOU = GetObject("LDAP://CN=" &ad_username& ",ou=Students,dc=bullitt,dc=ketsds,dc=net") 
		objOU.Physicaldeliveryofficename = NewPhyOffice
		objOU.SetInfo 

'Log The Move
LogEvent "66", ad_username&" has left "&RemoveFromADGroup&" and entered "&AddToADGroup&"."
'Delete Record from Move Table
RemoveRecordFromMoveTable rs.Fields.Item("id").Value
Else
LogEvent "66", ad_username&" can not be found in Active Directory."
End If
rs.MoveNext
Loop	
rs.Close
cn.Close
End Function
