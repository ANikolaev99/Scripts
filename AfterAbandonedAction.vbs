Dim OutboundCampaignManager
Dim Campaign
Dim oRequest
Dim request
Dim Requests
Dim PhoneNumber
Dim i
Dim itemsCount
Dim count
Dim Queue
Dim id
Dim Phone
Dim Phon
Dim ClientNameForOC
Dim ClientCodeForOC
Dim QueueDroppedCall
Dim tmp_phone_len
	count=0
		OutputDebugString "-----------------------START-----------"
	Set Queue = Session.ACDQueueItem.Queue
	Phone = Session.ACDQueueItem.Address
		QueueDroppedCall = Session.ACDQueueItem.Queue.DisplayName
	id = Queue.QueueId
		If Session.ACDQueueItem.Queue.QueueID = "{FC1BD996-58AB-4B89-A9A9-36A4EC8D7808}" then
		OutputDebugString "--------------------------------anytime Velcom	"
			Phon = 80&right(Phone,9)
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{D903A8F0-EC02-4782-916F-925EF32A8F43}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
						Set Request = oC.AddRequest(phone, customer_name, customer_name)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{55FE6090-2B30-41C5-A826-9B19CD38ECBB}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "-------------------------------- A-100 AZS MGTS"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{037E1E7F-7B24-47BD-8C09-CB578D2DBABF}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				elseIf Session.ACDQueueItem.Queue.QueueID = "{5EA597B8-E0C5-4E13-AA91-2BD3C19E2E1D}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "-------------------------------- A-100 AZS Velcom"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{037E1E7F-7B24-47BD-8C09-CB578D2DBABF}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				elseIf Session.ACDQueueItem.Queue.QueueID = "{5FB733B8-9F85-44E9-839C-42CB8BD02FF1}" then

			Phon = 80&right(Phone,9)
			OutputDebugString "---------------------A-100 AZS MTC"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{037E1E7F-7B24-47BD-8C09-CB578D2DBABF}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{216CD0EB-B185-404E-9F05-42E2EC378EFC}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "-------------------------------- A-100 DEV MGTS"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{037E1E7F-7B24-47BD-8C09-CB578D2DBABF}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{AADCE7A4-6330-4025-8182-74719219E086}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------A-100 DEV VELCOM"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{037E1E7F-7B24-47BD-8C09-CB578D2DBABF}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{9358C656-8042-4711-B212-6127839788EF}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------A-100 DEV MTC"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{037E1E7F-7B24-47BD-8C09-CB578D2DBABF}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{3335F71F-8A3E-4398-AB07-86E877143081}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------INVITRO MGTS"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0C6F7B93-CFA3-4B34-B47B-03852144471C}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
						'------
								request.NextAttemptDate = calls( 3, j )
'				call FileLogger("next call time (from DB): " & calls( 3, j) & "; request NextAttemptDate property: " & request.NextAttempDate)
				request.Property("ExternalProgram") = "chrome.exe" 'call external browser
				request.Property("ExternalProgramCommandLineArgs") = "http://192.168.19.191/index.php?action=Login&module=Users"  'pass CRM page to the browser
				request.ApplyChanges
						'-------
						
							if Phon=PhoneNumber then
								wscript.quit
								OutputDebugString "--------------------------------INVITRO MGTS qit"
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{953D7C6C-939D-403D-A2A9-8857488E15DA}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------INVITRO LIFE"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0C6F7B93-CFA3-4B34-B47B-03852144471C}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{A266DAF2-F06B-4B7A-AE4C-3AD021BDE182}" then
				Phon = right(Phone,11)
				OutputDebugString "--------------------------------ТД ПРАЙД"
				Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
				Set Campaign = oCM.GetCampaignByID ("{925D7EC6-9EF3-4C78-8BFD-EEDB2A4F57A0}")
				Set oRequest = Campaign.Requests
				itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,11)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{699AA989-91BE-4844-A8BE-6FBC74942805}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------INVITRO Velcom"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0C6F7B93-CFA3-4B34-B47B-03852144471C}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{C4FEE3A6-9028-426E-AF14-8BC008B4D066}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------INVITRO MTC"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0C6F7B93-CFA3-4B34-B47B-03852144471C}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if				
			elseIf Session.ACDQueueItem.Queue.QueueID = "{8D223829-E175-42CD-8605-89423ECC10D2}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------INVITRO Амадей Клиник"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0C6F7B93-CFA3-4B34-B47B-03852144471C}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{E55E17FA-98DF-4348-B05E-C961F4E002E4}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------INVITRO Страховая"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0C6F7B93-CFA3-4B34-B47B-03852144471C}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{EF900733-F170-4A6C-A5EC-A80ACFF25088}" then
				Phon = 80&right(Phone,9)
				OutputDebugString "--------------------------------WINGO 1"
				Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
				Set Campaign = oCM.GetCampaignByID ("{B5EB6E97-C650-4E0D-8F82-17D2F432311E}")
				Set oRequest = Campaign.Requests
				itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{0DF4079B-936C-4DC6-AD7F-B16BBAA0F510}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------WINGO 2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{B5EB6E97-C650-4E0D-8F82-17D2F432311E}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
		elseIf Session.ACDQueueItem.Queue.QueueID = "{1B24438B-FFBA-485E-B6EA-F156D1F4A03F}" then
				Phon = 80&right(Phone,9)
				OutputDebugString "--------------------------------------------WINGO 3"
				Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
				Set Campaign = oCM.GetCampaignByID ("{B5EB6E97-C650-4E0D-8F82-17D2F432311E}")
				Set oRequest = Campaign.Requests
				itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{3203EFEF-1537-46FF-84A9-BEB90EB31275}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Migom GoodZone Velcom"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{C1D4C38C-1F1C-4962-83B7-CBCC6D0F2D14}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{F5A8A001-AADC-4AFB-A812-9BCAFCE07786}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Migom GoodZone MTC"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{D0A43560-ABC2-45A5-A17A-8871B92303DD}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{D42EA920-8D2A-4DC0-90C5-053DB0967CEA}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "-------------------------------Тверь 74822452890"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			elseIf Session.ACDQueueItem.Queue.QueueID = "{29C18E91-D8A6-497A-91BD-00EB98A1E0B0}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------СПБ Общий 78126432030"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				'!-----'
elseIf Session.ACDQueueItem.Queue.QueueID = "{67E4EDBE-B1C9-4BD4-84A3-0FC5ED54733C}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------СПБ на сайте мл.рф 78003332612"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{CBDAFC97-0BB7-41F7-AC90-B1D0610A34BC}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Xiaomi MTS"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{A5EB617B-45BA-4829-9D2D-720ACB673280}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "-------------------------------Смоленск 74812339800"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{53CB0361-9EAA-4D33-A36F-1EA4BA0F09AC}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Самара 78463004058"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{F16F90F3-8604-4009-A9DD-3E60646EDF2B}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Сайт мл.рф и переадр с 84956432535"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{4D2B0CC4-9D1A-494A-9103-DB73E4A28735}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Рязань 74912434047"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{48DAEF0F-11AC-4EF7-9DB4-8232850A74DA}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Ростов-на-Дону 78633332006"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{A9B9F567-86ED-4C3F-9866-49A2AD0B1EB0}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Псков 78112331083"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{9AB41F59-888D-489A-9F58-9A9B8948ACD5}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Проключен на КЦ РФ 74996382873"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{B8DAB2B3-756C-458F-A95D-69BC60721EAF}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Пермь 73422554043"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{D946BAF2-82F3-416B-B27A-BE600C2E1C1D}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Омск 73812207325"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{3AFE740E-A2E8-43FA-9966-8236FDC15B36}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Новосибирск 73833830209"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{309921CD-5F30-497B-A61A-BD7DFB772B52}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Нижний Новгород 78312621240"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{760AB11C-44AE-4779-B7DE-57FD6A7602DA}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Мелкие площадки 74991102958"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{9D7081FE-5673-4268-B54E-5ACCEAF85A58}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "-------------------------------Красноярск 73912287005"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{DCEC4D13-C131-4688-8A72-718E08EAFA1B}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Краснодар 78612018620"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{A63793FA-FE2E-43AF-BEF2-2E165954F31E}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Кострома 74942641344"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{BFD5D9E5-1082-4B4F-9678-FF7A1EDD82FC}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Калуга 74842207073"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{0910AF3D-ED9E-42F2-9B0E-FB46760953B6}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Из рук в руки 74991103188"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{397BE516-0134-477B-AB68-ED957AA4E2C4}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Иваново 74932267838"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{383C1521-D3E9-4E4E-902D-F0DF61B64248}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Екатеринбург 73432478513"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{A838297F-3647-4B51-826B-1107665B5339}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Воронеж 74733003238"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{9353B4BF-DEE4-4EDF-86DE-FC06A8D10A29}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Волгоград 78442962150"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{7528FE1D-D4F5-492C-AE71-BA8A43486418}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "-------------------------------Владимир 74922494415"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{37CA9C2C-0B7C-4CD6-B9E0-8479E9F0FA81}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Владивосток 74232025035"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{F195DA6D-9A1F-43C5-847B-9AB1CA92398F}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Великий Новгород 78162637993"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{5C4048AC-4CC2-4F0C-AC9F-9B99E40E8439}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Брянск 74832770343"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{0803C39A-1226-4154-A3AE-8CF29992857D}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Авито СПБ 78126158096"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{57CE4F32-47B4-4F08-811F-6A773DAFED1D}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Авито МСК3 74991104171"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{B70A3CED-0F31-43A4-BA5E-473604396185}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Авито мск2 74994901556"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{08332087-92C6-42D6-8BCD-6AFFCB525E44}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Авито МСК2 74991102905"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
elseIf Session.ACDQueueItem.Queue.QueueID = "{BC8C6FA7-C275-4366-B324-89D3137760FA}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Авито мск1 74993224377"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{9B61FA53-C683-49F3-A2B6-42D79421F200}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------78432122260 Казань	"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if

elseIf Session.ACDQueueItem.Queue.QueueID = "{DBDA608B-8F01-4360-8B40-79E78E6F4061}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------ОПТОВЫЙ ОТДЕЛ РФ"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
elseIf Session.ACDQueueItem.Queue.QueueID = "{C832AD32-777F-403F-808A-4E49E43FA529}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------ОТДЕЛ ОЦЕНКИ"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{103AB39E-ED11-45B9-A096-566DA000F0CB}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
'!!----'
elseIf Session.ACDQueueItem.Queue.QueueID = "{0DD2E49C-DCE3-42B9-9868-9C309AC6E49F}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Колцентр РБ"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{616383A5-DDD4-4E9B-A3C1-4924F0610836}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Американсиких авто"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{DBE964DC-FA85-4A9F-9B16-14F1E64A4591}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Ауди 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{9602E696-07D3-43FE-8FFC-933D2973F621}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 БМВ 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{A52FA33F-2D9F-4067-B633-876F5908F68E}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Микроавтобусы 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{125E4D7C-A639-4F7A-BE8E-6FA5F8509BAD}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Микроавтобусы 2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{BC6A27DB-F49E-4DFF-8373-0DA0E9552D0D}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Форд 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{2F8D0D1B-10B3-4E2F-AF74-6AFDD40B23D2}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Форд 2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{6B78A2FD-D096-4E68-AA95-BC9ACDE17E69}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Французы1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{EED66010-EE97-4AD0-A82A-252308BC32D1}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Французы2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{CF9B66DF-465B-4BB1-9189-4D25CFD8EB7F}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Французcких авто"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{2FE3378C-DB3E-4D44-84AB-DB1A332CE8E4}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Форд"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{FC438A63-7A8F-42CC-A381-893268623469}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти для Японсиких авто"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{95449FAB-2F10-4ADC-8A7F-38F8E601C7E6}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Опель"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{1DD7AA41-13A1-48F3-B2A9-DD2BA14832FE}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Запчасти к Двигателям 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{23E5AE3F-982C-4776-B910-6857214B85E1}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Запчасти к Двигателям 2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{9F174F60-FA42-4C66-A8BF-70AA9B008FA8}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчатисти Вольво"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{2CAA64D7-BD84-41F5-85E2-E650967BFA2A}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти к Фольксваген"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{926761D7-009A-49DC-8E4E-268D73D8C754}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Мерседес 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{509195A4-D1F2-4C1A-B231-A1DB7AF7B96A}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Мицубиси_Ниссан 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{95A709C0-A420-43BB-AFEE-F0C5870BADBC}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Двигатели"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{E441FF16-FA98-4F38-9F4F-FE3AFF068AC5}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Двигатели1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{3B52BF58-A57C-4A33-83AF-5F551DB45E62}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Опель 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{343ADA2C-6043-4448-AEFD-91EBA3E8D03E}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Шины и диски"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{8236AE4E-88C2-490B-8176-9DFC3665836A}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Шины и диски1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{68369739-1B0E-41E2-BF5E-552AF6D74D2E}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Шины 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{6C80B001-29C5-4566-B588-312E5390BB4F}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад2.1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{BFDB320A-1A84-4AE1-BCF7-87D303428A51}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад2.2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{95823E8B-61AD-4A76-8136-6A3CB2F2EC92}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад2.3"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{95832B68-5666-493F-B73B-CF97F652B29D}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Двигатели 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{473F4B06-9EF4-4762-BBC7-6122F9CCEFE9}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Двигатели 2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{24183BB4-0713-437A-8031-CF0C5CF6E219}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Фольцваген 1"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{347E457B-FCBB-4558-8CC3-2ADDE873B2B8}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Опель"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{5BD137CB-473C-4FC0-BCA6-BAEEF25F9411}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Ауди"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{4E9C6DEF-17E9-493F-BDAC-A270906C4592}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Форд"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{D5F69489-11E2-4287-A95B-8AA6A3932B57}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Французы"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{3A4507C5-C58E-4626-AB24-3D293B7E60C6}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Мицубиси/Нисан"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{30AE093C-405C-4554-A8FC-4AB684B52D53}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Фольцваген"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{64B368EF-D058-43AC-BF11-63F37F184F65}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Шины"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{8736475C-128C-4648-A3A1-8EBC24D2AAC3}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 БМВ"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{49B5A735-DB8D-4115-BF22-278FC42B2570}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Бусы"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{0C4464DB-C2ED-4FA4-8669-FFD6FEB0D3FD}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 З/П двигатели"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{C832AD32-777F-403F-808A-4E49E43FA529}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------ОТДЕЛ ОЦЕНКИ"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{DD25C942-F23C-4531-9E59-27190D385BA5}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Двигатели"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{8791CC13-DC6F-446D-BA21-5339F1DD64CF}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Мерседес"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{B829D6C8-A51B-4049-B487-635A9A11853D}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Cклад2"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{DDD0B8D5-9528-4E5D-A654-A76C32CF5C66}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Cклад3 Запчасти Американских авто"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{3FA37B64-F515-4360-9214-3B8DCA511CC2}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Опель"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{0CEEE7D3-E715-4644-8A01-990268816ADD}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Японских авто"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{E69F155B-C656-4259-9A8E-B38FF54AD449}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Французких авто"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{F9C91631-B6EE-4C1A-8BBE-EB8E7E841DA0}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Фольцваген"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{675BAD05-0C70-4AD0-B03C-9EC7CDC18F09}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Запчасти Форд"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{42F61006-E866-47E9-A68D-9AEB9D2220BF}" then
			Phon = Right(Phone,9)
			OutputDebugString "-------------------------------Склад3 ДВС"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
				elseIf Session.ACDQueueItem.Queue.QueueID = "{AD12FD77-8961-4F62-B505-5837326C2085}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Шины"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
								elseIf Session.ACDQueueItem.Queue.QueueID = "{9D5054EC-2A89-4553-8D27-7E33A5C6184A}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад3 Вольво"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
				
								elseIf Session.ACDQueueItem.Queue.QueueID = "{C1E8D82D-4420-444A-BBF4-6AD9CBEA3D3D}" then
			Phon = Right(Phone,9)
			OutputDebugString "--------------------------------Склад1 Шины"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{0D96BE19-A3F9-41B3-9ED4-605D218B37DE}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			'!!!---'
			elseIf Session.ACDQueueItem.Queue.QueueID = "{C0FAC306-2A41-4459-974F-6F92C5E35655}" then
			Phon = 80&right(Phone,9)
			OutputDebugString "--------------------------------Входящие звонки"
			Set oCM = CreateObject("INFra.CC.OutboundCampaignManager")
			Set Campaign = oCM.GetCampaignByID ("{FE7FA528-61B4-416E-8576-6753EC384460}")
			Set oRequest = Campaign.Requests
			itemsCount = oRequest.Count
				if itemsCount > 0 then
					Requests = oRequest.GetItems (0, itemsCount)
						for i=0 to Ubound(Requests)
						Set request = Requests(i)
						PhoneNumber = 80&Right(request.PhoneNumber,9)
							if Phon=PhoneNumber then
								wscript.quit
							end if
						next
				end if
				if count=0 then
					Set Request = Campaign.AddRequest(Phon, "ClientCode", QueueDroppedCall)
				end if
			
				
		end if
		Request.ApplyChanges
		