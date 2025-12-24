<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 

<!--#include file="..\##GlobalFiles\DefaultColours.asp" --> 

<% 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

On Error Resume Next

Dim User
Dim ErrMsg
Dim ErrColour 
User = Session("UserName")

If Session("UserId") = Cint(4) Then
    ErrMsg = Session("SystemError")
    ErrColour = "#000000"
Else
    ErrMsg = "Admin have been informed"
    ErrColour = "#33CC33"
End If

Dim Emsg
                             
Set Emsg = Server.CreateOBject( "JMail.Message" )
Emsg.Logging = True
Emsg.silent = True

Emsg.priority = 2
Emsg.From = "no-reply@mediaco.co.uk"
Emsg.FromName = "HS Reporting"

If Session("Smtp") = "smtp.mail.yahoo.com" Then
    '## For testing at home
    Emsg.AddRecipient "fuftest-upload@yahoo.co.uk"    
ElseIf Session("Smtp") = "192.168.20.50" Then
    '## All error emails come to me 
    Emsg.AddRecipient"alan@mediaco.co.uk", "Alan Holgate"     
End If

Emsg.Subject = "HS Reporting Error"  
Emsg.Body = "There is a problem with the Web Site. " & vbCrLf & vbCrLf & Session("SystemError")
Emsg.Body = Emsg.Body & vbCrLf & "User was " & User

Emsg.Send( Session("Smtp") )
Set Emsg = nothing

Session.Abandon

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>System Error</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/ban.ico" type="image/x-icon" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
</head>

<body style="padding: 0px; margin: 0px">

<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
        <td align="left" valign="bottom" height="100">
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />
        </td>        
    </tr>
    <tr>
	    <td height="8" valign="top" >
	        <hr style="border-style: none;  height: 4px; background-color: <%=MplCyan%>; display: block;" />	        
	    </td>
    </tr>
    <tr>
        <td align="center"><font style="color: #0069AA; font-weight: bold; font-size: 16px;">Machine&nbsp;Fault&nbsp;log</font></td>
    </tr>        
</table>

<table style="width: 100%; position: absolute; top: 45%; padding-right: 20px; padding-left: 20px;" > 
    <tr>
        <td>
            <h2 align="center" style="color: #FF0000">System Error</h2> 
            <h3 align="center" style="color: #000000">An unrecoverable error has occured whilst processing data</h3>
            <h3 align="center" style="color: <%=ErrColour%> "><%=ErrMsg%></h3>          
        </td>
    </tr>    
</table>
<br />
<br />
<table style="width: 100%; position: absolute; bottom: 5px;  padding-right: 10px; padding-left: 10px;">
    <tr>  
        <td height="50" >
            <hr style="width: 98%; border-style: none;  height: 1px; background-color: <%=MplCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way,> 
            Trafford Park, Manchester M17 1BS</font></p>
        </td>
    </tr>
</table> 

</body>

</html>
