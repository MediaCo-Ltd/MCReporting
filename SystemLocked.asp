<%@Language="VBScript" Codepage="1252" EnableSessionState=False%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store" 

'On Error Resume Next

Dim TestFso
Dim Refresh
Dim MsgStatus
Dim MsgStatusStyle
Dim StatusIcon

Dim Fso
Set Fso = Server.CreateObject("Scripting.FileSystemObject")

'## Check if Sql Server is on line, check for existance of ClarityBackup folder
If Fso.FolderExists(ChkServer) = False Then
    Set Fso = Nothing
    'StatusIcon = "Images/Ban.ico"
    StatusIcon = "Images/PenPad64.png"   
    MsgStatus = "This web site is currently disabled"
    MsgStatusStyle = "color: #FF3300; text-align: center; font-size: 20px;"
    Refresh = "<meta http-equiv='refresh' content='600' />"        
Else
    Dim StatusRS		 
    Set StatusRS = Server.CreateObject("ADODB.Recordset")
    StatusRS.ActiveConnection = strConStatus
    StatusRS.Source = "SELECT ID, Locked, WebLocked from Status"
    StatusRS.CursorType = 0
    StatusRS.CursorLocation = 3
    StatusRS.LockType = 1
    StatusRS.Open
    Set StatusRS.ActiveConnection = Nothing

    If Err <> 0 Then
        StatusRS.Close
        Set StatusRS = Nothing 
        StatusIcon = "Images/PenPad64.png" 
        MsgStatus = "This web site is currently disabled"
        MsgStatusStyle = "color: #FF3300; text-align: center; font-size: 20px;"
        Refresh = "<meta http-equiv='refresh' content='600' />"
    Else
        If StatusRS("Locked") = True Or StatusRS("WebLocked") = True Then
            StatusRS.Close
            Set StatusRS = Nothing 
            StatusIcon = "Images/PenPad64.png" 
            MsgStatus = "This web site is currently disabled"
            MsgStatusStyle = "color: #FF3300; text-align: center; font-size: 20px;"
            Refresh = "<meta http-equiv='refresh' content='600' />"
        Else
            StatusRS.Close
            Set StatusRS = Nothing
            StatusIcon = "Images/PenPad64.png"
            MsgStatus = "This web site is now active again"
            MsgStatusStyle = "color: #2DB32D; text-align: center; font-size: 20px;"
            Refresh = ""
        End If
    End If
End If

Err.Clear

'## Page will refresh after 10 Mins, if status is back to normal will allow user to go back
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="<%=StatusIcon%>" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <%=Refresh%>
</head>

<body style="padding: 0px; margin: 0px">
    <table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
        <tr>
		    <td align="left" valign="bottom" height="100">
                <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
            </td>            
        </tr>
	    <tr>
		    <td height="8" valign="top">
		        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		    </td>		    
	    </tr>	
    </table>
    
    <table style="position: absolute; width: 100%; top: 40%;" >
        <tr>   
            <td valign="top" >
                <p style="<%=MsgStatusStyle%>"><%=MsgStatus%></p>                
            </td>
        </tr>
    </table>
                  
    <table style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
        <tr>  
            <td height="50" >
                <hr style="width: 98%; border-style: none;  height: 1px; background-color: <%=NewCyan%>; display: block;" />
                <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
                Trafford Park, Manchester M17 1BS</font></p>
            </td>
        </tr>
    </table>
</body>  
</html>