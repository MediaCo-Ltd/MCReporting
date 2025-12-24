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

Dim CentreLogo
Dim CentreAlt

CentreLogo = "Images/" & Session("CentreLogo") 
If Session("CentreLogo") = "blank.jpg" Then
    CentreAlt = ""
Else
    CentreAlt = Session("Client") & " Logo"
End If

'### get all reasons & then reason type, reason typoe will need to do a loop to set th epicked one

Dim ReasonRs

Set ReasonRs = Server.CreateObject("ADODB.Recordset")
ReasonRs.ActiveConnection = Session("ConnHSReports")
ReasonRs.Source = "SELECT * From Reasons Where (HasLogs = 0) Order By Description"
ReasonRs.CursorType = Application("adOpenForwardOnly")
ReasonRs.CursorLocation = Application("adUseClient")
ReasonRs.LockType = Application("adLockReadOnly")
ReasonRs.Open
Set ReasonRs.ActiveConnection = Nothing


Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"
End If

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
     <script type="text/javascript" >
        function Relocate() 
        {
            var Id = document.getElementById("hReasonId").value;
            window.location.replace('HsEditReason.asp?Id=' + Id);
        }                
    </script>
</head>

<body style="padding: 0px; margin: 0px" >           


<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100" colspan="3">
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
            &nbsp;<a href ="javascript:window.location.replace('HsAdmin.asp');" style="color:<%=NewBlue%>">Return&nbsp;to&nbsp;HS&nbsp;admin&nbsp;page</a>
        </td>
	    <td  align="center" width="34%">
	        <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >HS&nbsp;Reporting&nbsp;Select&nbsp;Reason</font>
	    </td>
	    <td align="right" width="33%">&nbsp;</td>
	</tr> 
	<tr>
	    <td colspan="3" align="left">
	        <br />
	        <font style="font-weight: bold; color:red; font-size: 16px;" >&nbsp;Only&nbsp;Reasons&nbsp;without&nbsp;logs&nbsp;can&nbsp;be&nbsp;edited</font> 
	    </td>
	</tr>   
</table>
                                                        
                                                        
<table class="NonDataTables" style="width: 100%; position: absolute; top: 50%;" >        
    <tr >
        <td  width="150px" ><font size="2">Select&nbsp;Reason&nbsp;</font></td>
        <td  >
            
            <select name="cboReason" id="cboReason" onchange="javascript:HSReasonSelect();" > 
            <option value="" >Select Reason</option>
            <%
            While Not ReasonRs.EOF
                Response.Write ("<option value='" & Trim(ReasonRs("Id")) & "'>" & Trim(ReasonRs("Description")) & "</option>)") & VbCrLf
                ReasonRs.MoveNext
            Wend
            
            ReasonRs.Close
            Set ReasonRs = Nothing
            %>
            </select>   
        </td> 
            
    </tr>            
    <tr>
        <td height="20px" colspan="2">&nbsp;</td>
    </tr> 
  
</table>
                                                        
<table style="width: 100%; position: absolute; bottom: 12%; padding-right: 20px; padding-left: 20px;">
        <tr>
            <td>&nbsp;&nbsp;</td>
            <td >
                &nbsp;&nbsp;<input id="btnSubmit" name="btnSubmit" type="button" value="Submit" onclick="Relocate();" disabled="disabled" />&nbsp;&nbsp;
		        <input id="btnReset" name="btnReset" onclick="javascript:ResetPage();" type="button" value=" Reset " />&nbsp;&nbsp;                   
                <input type="<%=InputType %>" name="hReasonId" id="hReasonId" value=""/>&nbsp;&nbsp;
            </td>
        </tr>          
    </table>    
 
</body>  
</html>

