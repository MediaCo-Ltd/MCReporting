<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

'## Check if System is locked
'Server.Execute "SystemLockedCheck.asp"
'If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

'If Session("UserName") = "" Then Response.Redirect ("CloseWindow.asp")

Dim BrowserType
BrowserType = Ucase(Request.ServerVariables("http_user_agent"))

If Instr(1,BrowserType,"FIREFOX",1) > 0 Then
    Session("Browser") = "Firefox" 
ElseIf Instr(1,BrowserType,"EDGE",1) > 0 Then
    Session("Browser") = "Chrome"        
ElseIf Instr(1,BrowserType,"CHROME",1) > 0 And Instr(1,BrowserType,"SAFARI",1) > 0 Then
    Session("Browser") = "Chrome"
ElseIf Instr(1,BrowserType,"SAFARI",1) > 0 Then
    Session("Browser") = "Safari"
Else
    Session("Browser") = "Unknown"
End If


Dim strImagePath    
strImagePath =  Request.QueryString("Path")
             
    
'If strImagePath = "" Then Response.Redirect "CloseWindow.asp"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
    <head>
        <title>Machine Faults</title>
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
        <link href="CSS/MachineFaultCss.css" rel="stylesheet" type="text/css" />
        <link href="CSS/MachineFaultExtraCss.css" rel="stylesheet" type="text/css" />
        <link rel="shortcut icon" href="Images/W-S48.png" type="image/x-icon" />
        <script type="text/javascript" src="JsFiles/MachineFaultJSFunc.js"></script>
    </head>
    <body onload="javascript:ShowImageLoadChk();">
         
        <center>
            <br /><br />
            <%Response.Write "<img alt='' style='color:#0068B3;' width='90%' src='" & "FaultImages/" & strImagePath & "' border='2'><br /><br />"%>
            <font size='2'><a href ="javascript:window.close();" style="color:<%=NewBlue%>;">Close Window </a></font>
        </center>
        <br />
    </body>
</html>
