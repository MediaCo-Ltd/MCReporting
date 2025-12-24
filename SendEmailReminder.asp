<%@language="vbscript" codepage="1252"%>
<%Option Explicit%>
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%

Dim IdToGet
IdToGet = Request.QueryString("Id")

If Isnumeric(IdToGet) = True Then

    '## Sends a password reminder email
    Dim msg
    Dim UsersRs

    Set UsersRs = Server.CreateObject("ADODB.Recordset")
    UsersRs.ActiveConnection = Session("ConnMcLogon")
    UsersRs.Source = "SELECT Id, Password, UserName, PwEmail FROM Users WHERE (Id =" &  IdToGet & ")"
    UsersRs.CursorType = Application("adOpenForwardOnly")
    UsersRs.CursorLocation = Application("adUseClient")
    UsersRs.LockType = Application("adLockReadOnly")
    UsersRs.Open

    If (Not UsersRs.EOF Or Not UsersRs.BOF) Then 
	    Set msg = Server.CreateOBject( "JMail.Message" )
	    msg.Logging = True
	    msg.silent = True
	    msg.From = "no-reply@mediaco.co.uk"
	    msg.FromName = "MediaCo Reporting"
    	
	    If Session("Smtp") = "mx496502.smtp-engine.com" Then        
            '## For testing at home
            msg.MailServerUserName = "warren.morris@mediaco.co.uk"
            msg.MailServerPassWord = "Q6)9l7Sit6dAV*"
        End If
        
	    msg.AddRecipient trim(UsersRs("PwEmail")), UsersRs("UserName")
	    msg.Subject = "Your login details" 
	    msg.Body = "Hello. " & UsersRs("UserName") & vbCrLf & vbCrLf 
	    msg.Body = msg.Body & "Your password is: " & UsersRs("Password").Value & vbCrLf & vbCrLf & "Thank you."	
    	
	    If Not msg.Send(Session("Smtp")) Then				
		    Session("LogonErrMsg") = "Unable to email, ask Alan or Dave for your PW"
	    Else
		    Session("LogonErrMsg") = "An email has been sent with your details."
	    End If
    Else
	    Session("LogonErrMsg") = "No email record has been found for you"
    End if

    UsersRs.Close
    Set UsersRs = Nothing
    
Else
    Session("LogonErrMsg") = "No email record has been found for you"
End If

Response.Redirect ("Login.asp")

'***************************************************************************************
%>
