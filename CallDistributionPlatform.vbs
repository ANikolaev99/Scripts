'---------------------------------------------------
'---------------------------------------------------
' by Koshelev Mastertelecom
' 2022-10-14
' Для колл-центра установлено рабочее время. В рамках этого времени для получения вызовов колл-центр должен присылать состояние параметров очереди звонков:
'	Формула = ("среднее_время_разговора_за_15_минут" * "количество_звонков_в_очереди" + 1) / "количество_залогиненных_агентов"
'---------------------------------------------------
'---------------------------------------------------

'---------------------------------------------------
'---------------------------------------------------
' Основные константы 
'---------------------------------------------------
'---------------------------------------------------

Dim MainQueueID : MainQueueID = "00A892E5-0BB2-49E4-8FAA-04629ED7E8FB" 'ID очереди по которой идет расчет
Dim centerId : centerId = "869f9131-984c-447e-b0ba-9ef32de0e75c" 'уникальный идентификатор колл-центра

'---------------------------------------------------
' Рабочее время КЦ
'---------------------------------------------------
Dim WorkHourStart : WorkHourStart = 9 'Время начала
Dim WorkHourStop : WorkHourStop = 21 'Время окончания

'---------------------------------------------------
'---------------------------------------------------
' Системные константы 
'---------------------------------------------------
'---------------------------------------------------
Dim LogFileName : LogFileName = "C:\Program Files\INFRATEL\integrations\Call-distribution-platform\Log.txt"
Dim LogType : LogType = 2 ' 0- win console, 1- debug console, 2- file, other values - log is off

'---------------------------------------------------
'---------------------------------------------------
' Данные подключения к SQL MigtyCall
'---------------------------------------------------
'---------------------------------------------------
Dim SQLDataSource : SQLDataSource = "MC2"
Dim SQLUserID : SQLUserID = "CDP_user"
Dim SQLPassword : SQLPassword = "qwerty-123"

'---------------------------------------------------
'---------------------------------------------------
' Сервисные методы
'---------------------------------------------------
'---------------------------------------------------

' Логирование сообщений
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


' Получаем среднее время разгоовора
Function FindAgent(Extension)

	Dim Queue, Agent, Connection
	Dim AgentLogin : AgentLogin=""
	
	Set Connection = WScript.CreateObject("INFRA.ACD.Connection")
	LogMessage( "Текущая ACD очередь = " + Session.ACDQueueItem.Queue.QueueId)
	Set Queue = Connection.GetQueueByID(Session.ACDQueueItem.Queue.QueueId)
		
	For Each Agent in Queue.Agents
		If Agent.Property("Extension")=Extension Then
			AgentLogin=Agent.LoginName
		End If
	Next
	
	LogMessage( "Логин найденного оператора = " + AgentLogin)
	FindAgent = AgentLogin

End Function

' Рассчитываем среднее время разгоовора за 15 минут из БД Migtycall 
Function SQLAVGTalkTime(QueueID)
	LogMessage( "Начинаем подключение к SQL... ")
	Dim cnnMain, RecordCall, AVGTalkTime
	'Подключаемся к БД  Migtycall
	Set cnnMain = CreateObject("ADODB.Connection")
	cnnMain.Provider = "SQLOLEDB"
	cnnMain.ConnectionString = "User ID=" & SQLUserID & ";Password=" & SQLPassword & ";Data Source = " & SQLDataSource & ";" & "Initial Catalog = INFRAVISOR"
	cnnMain.Open
	LogMessage( "Подключение к SQL прошло ")
	
	LogMessage( "SELECT AVG(Duration) FROM [INFRAVISOR].[dbo].[IV_CallRecord] where QueueID='"& QueueID &"' and StartTime>DATEADD(mi,-15,getdate())")
	Set RecordCall = cnnMain.Execute("SELECT isnull(AVG(Duration),0) FROM [INFRAVISOR].[dbo].[IV_CallRecord] where QueueID='"& QueueID &"' and StartTime>DATEADD(mi,-15,getdate())")
	If RecordCall.EOF or RecordCall.BOF then
			LogMessage( "Записей по условиям нет, AVGTalkTime=0")
			AVGTalkTime=0
			cnnMain.Close
	Else
	
		AVGTalkTime = RecordCall.Fields(0).Value
		LogMessage( "AVGTalkTime = " + CStr(AVGTalkTime) )
		cnnMain.Close
	
	End If
	
	SQLAVGTalkTime = AVGTalkTime
	
End Function

'Получаем кол-во агентов
Function AgentWork(QueueID)

	Dim acdConnection
	Dim supervisorSession

	Dim queue
	Dim agent


	Set acdConnection = CreateObject("INFra.ACD.Connection")
	Set supervisorSession = acdConnection.OpenSupervisorSession()

	Dim agentwrkn
	agentwrkn = 0
	For Each queue In supervisorSession.Queues
		if queue.queueid="{"& QueueID &"}" then
		'if queue.displayname="СДЭК" then
			For Each agent In queue.Agents
				if agent.Sessions.count>0 then
				'wscript.echo agent.state
					if agent.State=queue.AgentAvailableStatus then
						agentwrkn = agentwrkn + agent.Sessions.count
					end if
				end if
			Next
		end if
	Next

LogMessage("AgentWork = " + CStr(agentwrkn) )

AgentWork=agentwrkn

End Function

'Рассчитываем кол-во звонков в очереди
' Не работает на SP16 в этой конфигурации админ писал

'Function CountCallQueue(QueueID)
'
'	Dim acdConnection
'	Dim supervisorSession
'
'	Set acdConnection = CreateObject("INFra.ACD.Connection")
'	Set supervisorSession = acdConnection.OpenSupervisorSession()
'
'	Dim queue
'	Dim statistics
'    Dim valuesWithNames
'    Dim waiting
'	
'	For Each queue In supervisorSession.Queues
'		if queue.queueid="{"& QueueID &"}" then
'		'if queue.displayname="СДЭК" then
'			 Set statistics = queue.Statistics
'			valuesWithNames = statistics.ValuesWithNames
'			 waiting = valuesWithNames(18, 1)
'		end if
'	Next
'	
'	LogMessage( "waiting = " & CStr(waiting))
'	
'	CountCallQueue = waiting
'
'End Function

'Рассчитываем кол-во звонков в очереди
Function CountCallQueue(QueueID)

	Dim acd
	Dim oss
	Dim pos
	Set acd = CreateObject("infra.acd.connection")
	Set oss = acd.OpenSupervisorSession()
	Set pos = oss.CustomerSessions
	dim i: i = 0
	
	for each request in pos
		if request.Queue.QueueID = "{"& QueueID &"}" then
			i = I + 1
		end if
	next
	LogMessage( "waiting = " & CStr(i))
	
	CountCallQueue = i

End Function

' Получение токена
Function PutMetric(averageWaitTime, TokenL)
	LogMessage("Начинаем отправку данных...")
	Dim authCookie

	Dim objHTTP,requestBody,serviceURL',responseText

	sUrlAuth = "http://call-distribution-platform.cdek.ru/api/queue/metrics"
	
	requestBody = "{""centerId"":""" & centerId & """,""averageWaitTime"":" & averageWaitTime & "}"
	LogMessage(sUrlAuth + " :: " + CStr(requestBody))
	
	Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	objHTTP.open "POST", sUrlAuth, false
	objHTTP.setRequestHeader "X-Auth-Token", TokenL
	objHTTP.setRequestHeader "Content-type","application/json"
	objHTTP.send requestBody
	
	If objHTTP.status=200 Or objHTTP.status=204 Then
		LogMessage( "Данные отправлены Status: " + CStr(objHTTP.status))
		LogMessage( "Данные отправлены Text: " + CStr(objHTTP.responseText))
	Else
		LogMessage( "Ошибка отправки Status: " + CStr(objHTTP.status))
		LogMessage( "Ошибка отправки Text: " + CStr(objHTTP.responseText))
	End If

	Set objHTTP = nothing

End Function

'Забрать токен
Function GetToken ()

	Dim MyFileSystemObj, readFile, readline

	Set MyFileSystemObj = CreateObject("Scripting.FileSystemObject")
	Set readFile = MyFileSystemObj.OpenTextFile("C:\Program Files\INFRATEL\integrations\Call-distribution-platform\Token.txt")
	readline = readFile.ReadAll
	readFile.Close
	
	GetToken = readline

End Function

Function ProcessCall()
	
	If Hour(Now) >= WorkHourStart and Hour(Now) < WorkHourStop Then
	LogMessage( "Start")
	
	Dim Token, AVGTT, AW, CCQ, AWT, PM,AWTstr
	Token = GetToken ()
	LogMessage( "Token = " + CStr(Token))
	
	FOR i=0 to 30
		
		If (i Mod 15)=0 Then
			AVGTT=SQLAVGTalkTime(MainQueueID)
		End if
		
		AW=AgentWork(MainQueueID)
				
		CCQ=CountCallQueue(MainQueueID)
		
		If AW <> 0 Then 
				
			AWT = (AVGTT*CCQ+1)/AW
		
			LogMessage( "AWT = " & CStr(AWT))
		
			AWTstr = Replace(CStr(AWT), ",", ".")
				
			PM = PutMetric(AWTstr, Token)
		End If
		
		WScript.Sleep 2000
	NEXT
	End If
	
End Function

'-------------------------------------------------------
ProcessCall()
'-------------------------------------------------------

