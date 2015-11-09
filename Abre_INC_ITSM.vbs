'============================================================
'Autor: Alexandre
'Abertura de registros de Incidentes via VBScript - Out/2015
'============================================================

'On Error Resume Next

'Objeto para requisições WEB
Set objHTTP = CreateObject("MsXml2.ServerXmlHttp")
objHTTP.SetOption 2, 13056 'Ignora erros de certificados (https)

'--------------------------------
'Determina a URL WebService ITSM Remedy
'http://<midtier_server>/arsys/WSDL/public/<servername>/HPD_IncidentInterface_Create_WS
url = "http://<YOUR-URL>/arsys/services/ARService?server=<YOUR-APPSERVER>&webService=HPD_IncidentInterface_Create_WS"

'Usuário e senha para acesso ao ITSM Remedy
user = "user"
pass = "password"

'Monta XML na variavel dados
dados = "<SOAP-ENV:Envelope xmlns:ns0='http://schemas.xmlsoap.org/soap/envelope/' xmlns:ns1='urn:HPD_IncidentInterface_Create_WS' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:s0='urn:HPD_IncidentInterface_Create_WS' xmlns:SOAP-ENV='http://schemas.xmlsoap.org/soap/envelope/'>"&_
		"<SOAP-ENV:Header>"&_
			"<s0:AuthenticationInfo>"&_
				"<s0:userName>"&user&"</s0:userName>"&_
				"<s0:password>"&pass&"</s0:password>"&_
			"</s0:AuthenticationInfo>"&_
		"</SOAP-ENV:Header>"&_
		"<ns0:Body>"&_
			"<ns1:HelpDesk_Submit_Service>"&_
				"<ns1:Assigned_Group>Assigned_Group-here</ns1:Assigned_Group>"&_
				"<ns1:Assigned_Support_Company>Assigned_Support_Company-here</ns1:Assigned_Support_Company>"&_
				"<ns1:Assigned_Support_Organization>Assigned_Support_Organization-here</ns1:Assigned_Support_Organization>"&_
				"<ns1:Department>Department-here</ns1:Department>"&_
				"<ns1:First_Name>First_Name-here</ns1:First_Name>"&_
				"<ns1:Impact>4-Minor/Localized</ns1:Impact>"&_
				"<ns1:Last_Name>Last_Name-here</ns1:Last_Name>"&_
				"<ns1:Reported_Source>Systems Management</ns1:Reported_Source>"&_
				"<ns1:Service_Type>User Service Request</ns1:Service_Type>"&_
				"<ns1:Status>New</ns1:Status>"&_
				"<ns1:Action>CREATE</ns1:Action>"&_
				"<ns1:Summary>Summary-here</ns1:Summary>"&_
				"<ns1:Notes>Notes-here</ns1:Notes>"&_
				"<ns1:Urgency>1-Critical</ns1:Urgency>"&_
				"<ns1:Work_Info_Locked>No</ns1:Work_Info_Locked>"&_
				"<ns1:Work_Info_View_Access>Internal</ns1:Work_Info_View_Access>"&_
				"<ns1:HPD_CI_ReconID>ReconID-here</ns1:HPD_CI_ReconID>"&_
			"</ns1:HelpDesk_Submit_Service>"&_
		"</ns0:Body>"&_
	"</SOAP-ENV:Envelope>"

'Monta a requisição URL
objHTTP.Open "POST", url, FALSE
'Monta Header
objHTTP.setRequestHeader "Soapaction", "urn:HPD_IncidentInterface_Create_WS/HelpDesk_Submit_Service"
objHTTP.setRequestHeader "Host", "YOUR-DOMAIN"
objHTTP.setRequestHeader "Connection", "close"
objHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
'Envia os dados
objHTTP.Send dados

'Guarda o retorno (codigo-fonte da página retornada) na variavel resp
resp = objHTTP.ResponseText
'--------------------------------

'Mostra popup com o retorno
msgbox resp

'==================== Se necessario habilitar para gravar log em arquivo) ==================================
'set arquivo = CreateObject("Scripting.FileSystemObject")
'Set arq = arquivo.OpenTextFile ("C:\log_itsm.txt",8,true,tristateusedefault)
'arq.writeline now()&" => "&resp
'arq.writeline dados
'arq.Close

'Set arq = Nothing
'set arquivo = Nothing
'==================== Se necessario gravar log em arquivo) ==================================

'msgbox "Fim"