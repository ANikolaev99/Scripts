'---------------------------------------------------
'---------------------------------------------------
' by Koshelev Mastertelecom
' 2022-10-14
'---------------------------------------------------
'---------------------------------------------------

'---------------------------------------------------
'---------------------------------------------------
' �������� ��������� 
'---------------------------------------------------
'---------------------------------------------------

Dim SpAnLogin : SpAnLogin = "Utel_VZ"
Dim SpAnPassword : SpAnPassword = "bc8d7020fcb3a4f3bec81905ad7f0a7d" ' md5 ��� ������ � hex-������� B3vbiol%

'---------------------------------------------------
'---------------------------------------------------
' ��������� ��������� 
'---------------------------------------------------
'---------------------------------------------------
Dim LogFileName : LogFileName = "C:\Program Files\INFRATEL\integrations\Call-distribution-platform\Log.txt"
Dim LogType : LogType = 2 ' 0- win console, 1- debug console, 2- file, other values - log is off

'---------------------------------------------------
'---------------------------------------------------
' ��������� ������
'---------------------------------------------------
'---------------------------------------------------

' ����������� ���������
Sub LogMessage(Text)
	Select Case LogType
		Case 0 ' win console
			WScript.Echo Text
		Case 1 ' debug console
			OutputDebugString "[dispatcher.vbs]: " & Text & vbCrLf, true
		Case 2 ' file
			On Error Resume Next
			Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
			Dim objLogFile : Set objLogFile = objFSO.OpenTextFile(LogFileName, 8, True)
			objLogfile.WriteLine Now & " : " & Text
			objLogFile.Close 
			On Error Goto 0
	End Select
End Sub


' ��������� ������
Function Authenticate()
	LogMessage("�����������...")
	Dim authCookie

	Dim objHTTP,requestBody,serviceURL',responseText

	sUrlAuth = "http://auth.api.cdek.ru/web/simpleauth/authorize"
	
	requestBody = "{""user"":""" + SpAnLogin + """,""hashedPass"":""" + SpAnPassword + """}"
	LogMessage(sUrlAuth + " :: " + requestBody)
	
	Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	objHTTP.open "POST", sUrlAuth, false
	objHTTP.setRequestHeader "Content-type","application/json"
	objHTTP.setRequestHeader "X-User-Lang","rus"
	objHTTP.send requestBody
	
	If objHTTP.status=200 Or objHTTP.status=204 Then
		responseText = objHTTP.responseText
		responseText = Replace(responseText, "{""token"":""", "")
		responseText = Replace(responseText, """}", "")
		LogMessage( responseText)

	Authenticate = responseText
	Else
		LogMessage( "������ ���������� Status: " + CStr(objHTTP.status))
		LogMessage( "������ ���������� Text: " + CStr(objHTTP.responseText))
		responseText = "000"
	End If

	Set objHTTP = nothing

	Authenticate = responseText
End Function

Function ProcessGetToken()
	LogMessage( "�������� �����")
	
	Dim Token
	Token = Authenticate()
	If (Token<>"000") Then
		Set objFS = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFS.CreateTextFile("C:\Program Files\INFRATEL\integrations\Call-distribution-platform\Token.txt", True)
		objFile.Write (Token)
		objFile.Close
		LogMessage( "����� ������� � ����")
	Else
		LogMessage( "�������� ����� �� �������")
	End If
	
	
End Function

'-------------------------------------------------------
ProcessGetToken()
'-------------------------------------------------------

