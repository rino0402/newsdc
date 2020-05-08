Option Explicit
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
End Function
Call Include("const.vbs")

Call Main()

Function usage()
    Wscript.Echo "idoreki.vbs [/db:DbName] [/test] [/skip:num] [/limit:num]"
    Wscript.Echo "<��>"
    Wscript.Echo "idoreki.vbs /db:newsdc-ono"
End Function

Function Main()
	dim	strArg
	dim	objDb

	'���O�����I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.UnNamed
		call usage()
		Exit Function
	next
	'���O�t���I�v�V�����`�F�b�N
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "db"
		case "limit"
		case "test"
		case "skip"
		case else
			call usage()
			exit function
		end select
	next
	'�f�[�^�x�[�X�I�[�v��
	set objDb = OpenDb()
	' �e�[�u�����R�[�h�Z�b�g
	dim	rsTable
	Set rsTable = Wscript.CreateObject("ADODB.Recordset")
'	rsTable.Open "idoreki", objDb, adOpenForwardOnly, adLockBatchOptimistic	,adCmdTableDirect
	rsTable.Open "idoreki", objDb, adOpenForwardOnly, adLockReadOnly		,adCmdTableDirect

	dim	lngCnt
	lngCnt = 0
	Do While rsTable.EOF = False
		lngCnt	= lngCnt + 1
		if LimitCheck(lngCnt) then
			exit do
		end if
		EchoMsg Right(space(10) & lngCnt,10) _
			  & " " & rsTable.Fields("JITU_DT") _
			  & " " & rsTable.Fields("JITU_TM") _
			  & " " & rsTable.Fields("JGYOBU") _
			  & " " & rsTable.Fields("NAIGAI") _
			  & " " & rsTable.Fields("HIN_GAI") _
			  & " " & rsTable.Fields("RIRK_ID") _
			  & " " & rsTable.Fields("RIRK_NAME") _
			  & ""
		rsTable.MoveNext
	Loop

	' �e�[�u��Close
	rsTable.Close
	set rsTable = Nothing
	'�f�[�^�x�[�X�N���[�Y
	set objDb = CloseDb(objDb)

End Function

Function OpenDb()
	dim	objDb
	dim	strDbName

	strDbName = GetOption("Db","newsdc")

	Set objDb = Wscript.CreateObject("ADODB.Connection")

	EchoMsg "�f�[�^�x�[�X�I�[�v��:" & strDbName
	objDb.Open strDbName

	Set OpenDb = objDb
End Function

Function CloseDb(byval objDb)

	EchoMsg "�f�[�^�x�[�X�N���[�Y:" & objDb.Properties("Data Source")
	objDb.Close

	set objDb = Nothing

	Set CloseDb = objDb
End Function
Sub EchoMsg(byval strMsg)
	Wscript.Echo strMsg
End Sub

Function GetOption(byval strName _
				  ,byval strDefault _
				  )
	dim	strValue

	strValue = strDefault
	if WScript.Arguments.Named.Exists(strName) then
		strValue = WScript.Arguments.Named(strName)
	end if
	GetOption = strValue
End Function
Function LimitCheck(byval lngCnt)
	dim	lngLimit
	dim	bLimit

	bLimit = False
	lngLimit = CLng(GetOption("limit",0))
	if lngLimit > 0 then
		if lngCnt > lngLimit then
			bLimit = True
		end if
	end if
	LimitCheck = bLimit
End Function
