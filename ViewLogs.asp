<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%

On Error Resume Next

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"


Dim LogDate
Dim LogFile 
LogDate = Replace(Date,"/","-",1,-1,1)

LogFile = "C:\Web Sites\MC Reporting\Logs\Session log " & LogDate & ".txt"


Dim fso
Dim txtfile
Dim Data 
Set fso = Server.CreateObject("Scripting.FileSystemObject")    
Set txtfile = fso.OpenTextFile(LogFile , 1, False)
Data = txtfile.ReadAll 
txtfile.close
Set txtfile = nothing
Set fso = nothing

Data = Replace(Data,VbCrLf,"<br />",1,-1,1)

If Data = "" Then Data = "No logs for " & LogDate

Session.Abandon

Err.Clear        

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/PenPad64.png" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
</head>

<body style="padding: 0px; margin: 0px" >
    <table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
        <tr>
		    <td align="left" valign="bottom" height="100">
                <img align="left" alt="logo" src='Images/mediaco_logo.jpg' width="160" />
            </td>            
        </tr>
	    <tr>
		    <td height="8" valign="top">
		        <hr style="border-style: none;  height: 4px; background-color: <%=MplCyan%>; display: block;" />
		    </td>
	    </tr>	
        <tr>
            <td height="20px"  valign="bottom" align="center">        <!-- color #0069AA; Blue--> 
                &nbsp;&nbsp;<font style="font-weight: bold; color:<%=NewBlue%>;" size="3">MediaCo Reporting Session Logs</font>
            </td>
        </tr>
    </table>
    
   
    
    <table style="width: 100%; position: absolute; top: 20%; padding-right: 20px; padding-left: 20px;" >            
        <tr>   
            <td valign="middle" width="33%" >
            <!-- colour #0069AA New blue or cyan #07AFEE  -->      
                <p><font size="2" color="<%=NewBlue%>">                                       
                    
                    <%=Data%>
                </font>
                </p>
            </td>
       
        </tr> 
        
    </table>
                
    
</body>  
</html>

