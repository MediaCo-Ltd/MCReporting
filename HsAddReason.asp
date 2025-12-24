<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

If Session("ConnHSReports") = "" Then Response.Redirect("Admin.asp")

Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"
End If

Dim ReasonRs
Dim ReasonList

Set ReasonRs = Server.CreateObject("ADODB.Recordset")
ReasonRs.ActiveConnection = Session("ConnHSReports")
ReasonRs.Source = "SELECT * From Reasons Order By Description"
ReasonRs.CursorType = Application("adOpenForwardOnly")
ReasonRs.CursorLocation = Application("adUseClient")
ReasonRs.LockType = Application("adLockReadOnly")
ReasonRs.Open
Set ReasonRs.ActiveConnection = Nothing

While Not ReasonRs.EOF
    If ReasonList = "" Then    
        ReasonList = Ucase(ReasonRs("Description"))
    Else
        ReasonList = ReasonList & "#" & Ucase(ReasonRs("Description"))
    End If
    ReasonRs.MoveNext
Wend

ReasonRs.Close
Set ReasonRs = Nothing


'## No idea what below is for, never set up

'## Use multi pick dropdown, showing deselected locations at side
'## List will be locations not to show as the saved data will be less
'## combo default will be show all = "" rather than select location
'## in all cases if  ExcludedLocations in Reasons = "" clause is ignored

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>HS Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/HSReportsJSFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px" >           

<form action="HsUpdateReason.asp" method="post" name="frmAddReason" id="frmAddReason"  onsubmit="return HSValidateReason('Add');" >
<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100"> 
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" /> 
        </td>           
    </tr>
    <tr>
	    <td height="8" valign="top" colspan="3">
	        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
	    </td>
    </tr>
    <tr>
        <td align="left" valign="top" width="33%">
            &nbsp;<a href ="javascript:window.location.replace('HsAdmin.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;HS&nbsp;admin&nbsp;page</a>
        </td>
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >HS&nbsp;Reporting&nbsp;Add&nbsp;Reason</font>
        </td>
        <td align="right" width="33%">&nbsp;</td>
     </tr>	
</table>
                                                        
                                                        
<table class="NonDataTables" style="width: 100%; position: absolute; top: 40%;" >        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Reason</font></td>
            <td valign="middle" width="300px">
            <input name="txtReasonID" type="text" id="txtReasonID" class="inputboxes" onfocus='this.select();' onmouseup='return false;' 
                onkeypress='return DisableEnterKey(event);' value="" maxlength="48" />       
            </td>
            <td align="left">&nbsp;&nbsp</td>                
        </tr>            
                
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>          
        <tr >                                                           
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Reason&nbsp;Type</font></td>
            <td align="left" ><!--style="visibility: hidden"-->
                <select id="cboReasonType" name="cboReasonType" onchange="javascript:HSReasonTypeSelect();"><!--  -->
                    <option value="">Select&nbsp;Reason&nbsp;Type</option>
                    <option value="1">Accident</option>
                    <option value="2">Incident</option>
                    <option value="3">Unsafe</option>
                 </select>
            </td>
            <td align="left" >&nbsp;                
            </td>                
        </tr>
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Active</font></td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkActive" name="chkActive" checked="checked" />       
            </td>
            <td align="right">&nbsp;</td>                
        </tr> 
</table>
                                                        
<table style="width: 100%; position: absolute; bottom: 12%; padding-right: 20px; padding-left: 20px;">
        <tr>
            <td width="10%">&nbsp;
                <%
                    If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                        Response.Write ("Don't Save Data&nbsp;&nbsp;")
                        Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                    End If          
                 %>            
            </td>
            <td width="10%">
                <input id="btnSubmit" name="btnSubmit" type="submit" value="Update" disabled="disabled" />                   
            </td>
            <td width="10%">
                <input type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            </td>
            <td >
                
                <%If InputType = "text" Then response.Write "frmName"%>
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmAddReason" />
                <%If InputType = "text" Then response.Write "hReasonList"%>
                <input type="<%=InputType %>" name="hReasonList" id="hReasonList" value="<%=ReasonList%>"/>
                <%If InputType = "text" Then response.Write "hReasonType"%>
                <input type="<%=InputType %>" name="hReasonType" id="hReasonType" value=""/>
                <%If InputType = "text" Then response.Write "hReasonChange"%>
                <input type="<%=InputType %>" name="hReasonChange" id="hReasonChange" value="1"/>
            </td>            
        </tr>          
    </table>    
    
</form>   
</body>  
</html>


