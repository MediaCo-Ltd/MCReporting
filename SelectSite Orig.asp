<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%

Session("Location") = Location 
Session("ClarityConn") = strConnClarity
Session("ConnStatus") = strConStatus

If Session("UserName") = "" Then Response.Redirect ("Login.asp")

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

Response.expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

Dim HSvisible
HSvisible = ""
If Session("ShowHS") = Cbool(False) Then HSvisible = " style='visibility: hidden' "

Dim HideRedoRs
Set HideRedoRs = Server.CreateObject("ADODB.Recordset")
HideRedoRs.ActiveConnection = Session("ConnStatus")    
HideRedoRs.Source = "SELECT IndoorRedoClientOnly FROM Status"
HideRedoRs.CursorType = Application("adOpenForwardOnly")
HideRedoRs.CursorLocation = Application("adUseClient")
HideRedoRs.LockType = Application("adLockReadOnly")
HideRedoRs.Open
Set HideRedoRs.ActiveConnection = Nothing

Session("HideIndoorRedo") = Not(HideRedoRs("IndoorRedoClientOnly"))

HideRedoRs.Close
Set HideRedoRs = Nothing

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/PenPad64.png" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" >
        function Relocate(strUrl)
        {
            document.body.style.cursor = 'wait';
            document.getElementById('loading').innerHTML = "Loading Data, Please wait...";
            window.location.replace(strUrl);
        }

        function RelocateNoMsg(strUrl)
        {
            window.location.replace(strUrl);
        }

        function LogOut() 
        {
            window.location.replace("LogOff.asp");
        }   
    </script>       
</head>

<body style="padding: 0px; margin: 0px">
    <table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
        <tr>
		    <td align="left" valign="bottom" height="100" colspan="3">
                <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
            </td>            
        </tr>
	    <tr>
		    <td height="8" valign="top" colspan="3">
		        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		    </td>
	    </tr>	
        
        <tr>
	        <td width="33%">&nbsp;&nbsp;<a id="logoffL" href ="javascript:LogOut();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a></td>
            <td align="center" width="34%">
                <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
                <font style="color: #0069AA; font-weight: bold; font-size: 16px;">MediaCo&nbsp;Reporting&nbsp;(<%=Session("UserName")%>)</font>
            </td>
            <td align="right" width="33%"><a id="logoffR" href ="javascript:LogOut();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
        </tr>
        
    </table>
    
    <table style="width: 100%; position: absolute; top: 30%; padding-right: 10px; padding-left: 10px;" >            
        <tr>   
            <td valign="middle" width="33%" >      
                <p><font size="3" color="<%=NewBlue%>">
                    <br  /><br  />                                  
                    <label <%=HSvisible %> onclick="javascript:RelocateNoMsg('HsSelectOption.asp');">
                        &nbsp;<img alt="HS" src="Images/Plus-icon.png" style="width: 20px; height: 20px"  />                        
                        &nbsp;HS Reporting                        
                    </label>
                    <%
                        If Session("ShowNC") = Cbool(True) Then 
                            Response.Write "<br /><br /><br />"
                            Response.Write "<label onclick='javascript:RelocateNoMsg(""NcSelectOption.asp"");'>"
                            Response.Write "&nbsp;<img alt='HS' src='Images/warning.png' style='width: 20px; height: 20px' />"
                            Response.Write "&nbsp;Non Conformance"
                            Response.Write "</label>"
                        End If
                    %>
                    
                    <%
                        If Session("ShowMF") = Cbool(True) Then
                            Response.Write "<br /><br /><br />"
                            Response.Write "<label onclick='javascript:RelocateNoMsg(""MFSelectOption.asp"");'>"
                            Response.Write "&nbsp;<img alt='MF' src='Images/W-S48.png' style='width: 20px; height: 20px' />"
                            Response.Write "&nbsp;Machine Faults"
                            Response.Write "</label>"
                        End If
                            
                    %>
                    
                    
                    <%
                        If Session("HideIndoorRedo") = Cbool(True) And Session("AddRedo") = Cbool(True) Then 
                            Response.Write "<br /><br /><br />"
                            Response.Write "<label title='Sytsem not active yet'>"
                            Response.Write "&nbsp;<img alt='Redo' src='Images/X1.ico' style='width: 20px; height: 20px' />"
                            Response.Write "&nbsp;In House Redo"
                            Response.Write "</label>"  
                        ElseIf Session("HideIndoorRedo") = Cbool(True) And Session("ViewRedo") = Cbool(True) Then 
                            Response.Write "<br /><br /><br />"
                            Response.Write "<label title='Sytsem not active yet'>"
                            Response.Write "&nbsp;<img alt='Redo' src='Images/X1.ico' style='width: 20px; height: 20px' />"
                            Response.Write "&nbsp;In House Redo"
                            Response.Write "</label>"                                              
                        Else
                            If Session("AddRedo") = Cbool(True) Or Session("ViewRedo") = Cbool(True) Then
                                Response.Write "<br /><br /><br />"
                                Response.Write "<label onclick='javascript:RelocateNoMsg(""IhrJobNo.asp"");'>"
                                Response.Write "&nbsp;<img alt='Redo' src='Images/X1.ico' style='width: 20px; height: 20px' />"
                                Response.Write "&nbsp;In House Redo"
                                Response.Write "</label>"
                            End If
                        End If
                    %>
                    
                   
                                                                            
                </font>
                </p>
            </td>
            <td valign="middle" align="center" width="33%" id="loading" style="font-size: medium; color: <%=NewBlue%>; font-weight: bold;">
                <noscript style="color:Red" >Your Browser has Javascript disabled<br />Please enable, to allow full functionality</noscript>                       
            </td>
            <td width="33%">&nbsp;</td>
        </tr>
        <tr>
            <td colspan="3">
            <%
            If Session("PC-Name") = "Home" or Session("UserId") = 4 Then
                Response.Write "ConnHSReports = " & Session("ConnHSReports") & "<br />"
                Response.Write "ConnNCReports = " & Session("ConnNCReports") & "<br />"
                Response.Write "ConnMachinefaults = " & Session("ConnMachinefaults") & "<br />"
                Response.Write "ConnMcLogon = " & Session("ConnMcLogon") & "<br />"
                Response.Write "ConnQC = " & Session("ConnQC") & "<br />"
                Response.Write "QcPath = " & Session("QcPath") & "<br />"         
                Response.Write "ClarityConn = " & Session("ClarityConn") & "<br />"
                Response.Write "StatusConn = " & Session("ConnStatus") & "<br />"
                Response.Write "RootPath = " & Session("RootPath") & "<br />"
                Response.Write "MfImagePath = " & Session("MfImagePath") & "<br />"
                Response.Write "HsImagePath = " & Session("HsImagePath") & "<br />"
                Response.Write "Location = " & Session("Location") & "<br />"
                Response.Write "Session(UserName)&nbsp;" & Session("UserName") & "<br />"
                Response.Write "Session(UserId) = " & Session("UserId") & "<br />"
                Response.Write "Session(PC-Name) = " & Session("PC-Name") & "<br />"
                Response.Write "Session(HideIndoorRedo) = " & Session("HideIndoorRedo")               
            End If  
                        
            %>
            </td>
        </tr>     
    </table>
    
    
                  
    <table  style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
        <tr>  
            <td height="50" >
                <hr style="border-style: none; height: 1px; background-color: <%=NewCyan%>; display: block;" />
                <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
                Trafford Park, Manchester M17 1BS</font></p>
            </td>
        </tr>
    </table>
</body>  
</html>