On Error Resume Next

Const ADS_SCOPE_SUBTREE = 10

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

objCommand.CommandText = _
"SELECT ADsPath FROM 'LDAP://ou=xxx,ou=xxx,dc=sdch,dc=local' WHERE objectCategory='user'"
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	strPath = objRecordSet.Fields("ADsPath").Value
	Set objUser = GetObject(strPath)

	Const E_ADS_PROPERTY_NOT_FOUND = &h8000500D

	passwordLastChanged = objUser.passwordLastChanged

	If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
		passwordLastChanged = "パスワード変更履歴なし"
	End If

	Wscript.Echo objUser.samaccountname &","&objUser.description &","&objUser.whenCreated & " GMT,"&passwordLastChanged

	objRecordSet.MoveNext
Loop
