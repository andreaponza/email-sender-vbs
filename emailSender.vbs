' Purpose: Alert of when a computer has been shutdown or rebooted

' Set Variables
Currentdate = Date()
CurrentTime = Time()
Set oShell = CreateObject("WScript.Shell")
strComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "Notice of svstem shutdown from " & strComputerName
objMessage.From = """Sender Name"" <myMail@mail.com>" 
objMessage.To = "destinaionMail@mail.com" 
objMessage.TextBody = "The system has been shutdown or rebooted on " & CurrentDate & " at " & CurrentTime & "."

'==This section provides the configuration information for the remote SMTP server.
 Set objConfig = objMessage.Configuration

objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

'Name or IP of Remote SMTP Server
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.server.com"

'Type of authentication, NONE, Basic (Base64 encoded), NTLM
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

'Your UserID on the SMTP server
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "userID"

'Your password on the SMTP server
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Password"

'Server port (typically 25)
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 

'Use SSL for the connection (False or True)
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP   server)
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objConfig.Fields.Update

'==End remote SMTP server configuration section==  

objMessage.Send
