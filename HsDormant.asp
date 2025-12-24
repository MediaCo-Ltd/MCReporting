<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<% 
'on error resume next

If Session("ConnMcLogon") = "" Then Response.Redirect("Admin.asp")

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

Dim DormantCount
Dim DormantRs
Dim DormantSql
Dim CboDisabled
Dim StatusBg
Dim StatusText
Dim LoggedByName
Dim DormantArray


DormantSql = "Select Logs.Id, Logs.Dormant, Logs.CreatedByName, Logs.CreatedDate,"
DormantSql = DormantSql & " Logs.Severity, Reasons.Description" 
DormantSql = DormantSql & " From Logs INNER JOIN Reasons ON Logs.GroupSelection = Reasons.Id"
DormantSql = DormantSql & " Where (Dormant = 1) Order By Id"

Set DormantRs = Server.CreateObject("ADODB.Recordset")
DormantRs.ActiveConnection = Session("ConnHSReports")
DormantRs.Source = DormantSql
DormantRs.CursorType = Application("adOpenForwardOnly")
DormantRs.CursorLocation = Application("adUseClient") 
DormantRs.LockType = Application("adLockReadOnly")
DormantRs.Open

If DormantRs.BOF = True Or DormantRs.EOF = True Then
    CboDisabled = " disabled='disabled' "
Else
    CboDisabled = ""
    ReDim DormantArray(DormantRs.RecordCount,4)
    While Not DormantRs.EOF
        DormantArray(DormantRs.AbsolutePosition,0) = DormantRs("Id")
        DormantArray(DormantRs.AbsolutePosition,1) = Cdate(DormantRs("CreatedDate"))
        DormantArray(DormantRs.AbsolutePosition,2) = DormantRs("Severity")
        DormantArray(DormantRs.AbsolutePosition,3) = Trim(DormantRs("Description"))
        DormantArray(DormantRs.AbsolutePosition,4) = Trim(DormantRs("CreatedByName"))
        DormantRs.MoveNext
    Wend
        
End If    

DormantRs.Close
Set DormantRs = Nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>HS Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/HSReportsAjaxFunc.js"></script>
    <script type="text/javascript" >
        function GetId(e, objId) 
        {
            var upkey;

            if (window.event)
                upkey = window.event.keyCode;     //IE
            else
                upkey = e.which;     //firefox

            if (upkey == 13) 
            {
                if (objId == 'SetDormant') 
                {
                    var RefId = document.getElementById(objId).value;
                    SetHsDormant(RefId);

                }
                else
                { document.getElementById(objId).focus(); }
            }
        }

        function GetCboId() 
        {
            var CboId = document.getElementById("cboSetVisible").value;
            if (document.getElementById("cboSetVisible").value != '') 
            {
                UnSetHsDormant(CboId);
            }            
        }             
    </script>     
</head>

<body id="admin" style="padding: 0px; margin: 0px">

<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
        <td align="left" valign="bottom" height="100" colspan="3">
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />
        </td> 
              
    </tr>
	<tr>
		<td height="8" valign="top" colspan="3">
		    <hr style="border-style: none;  height: 4px; background-color: #07AFEE; display: block;" />
		</td>
	</tr>
	
	<tr>
	    <td align="left" valign="top" width="33%">            
            &nbsp;<a href ="javascript:window.location.replace('HsAdmin.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;HS&nbsp;Admin&nbsp;</a>
        </td>
	    <td  align="center" width="34%">
	        <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;">HS&nbsp;Reporting&nbsp;Admin</font>
	    </td>
	    <td  align="right" width="33%">
	        &nbsp;
	    </td>
	
	</tr>		
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 20%;" cellpadding="0" cellspacing="0">
    <tr>   
        <td valign="middle" colspan="5" >
            <!-- colour #0069AA New blue or cyan #07AFEE  -->     
            <p>
            <font size="3" color="<%=NewBlue%>">To Make record dormant. Enter number and press enter</font>
            </p>
        </td>
    </tr>
    
    <tr>        
        <td colspan="5"><input type="text" id="SetDormant"  onmouseup='return false;'  onkeyup="GetId(event,'SetDormant');"  /></td>    
    </tr>
    
    <tr>
        <td height="15px" colspan="5">&nbsp;</td>
    </tr>

    <tr>   
        <td valign="middle" colspan="5">
            <!-- colour #0069AA New blue or cyan #07AFEE  -->     
            <p>
            <font size="3" color="<%=NewBlue%>">To Make record visible, select from dropdown. Hidden records if any will show below dropdown</font>
            </p>
        </td>
    </tr>
    
    <tr>        
            <td colspan="5">
                <select id="cboSetVisible" <%=CboDisabled%> onchange="javascript:GetCboId();">
                    <option value="">Select Record</option>
                    <%
                    If CboDisabled = "" Then
                        For DormantCount = 1 To UBound(DormantArray)
                            Response.Write "<option value=" & DormantArray(DormantCount,0) & ">" & DormantArray(DormantCount,0) & "</option>"
                        Next
                    End If
                    %>                
                </select>            
            </td>        
    </tr> 
    
    <tr>
        <td height="35px" colspan="5" align="center"  >
            <p><font size="3" color="<%=NewBlue%>">Hidden records</font></p>
        </td>
    </tr>       
       
    <tr >
        <th class="styleTHleft" width="5%" >Record Id</th>
        <th class="styleTHstd" width="10%">Created</th>	        
        <th class="styleTHstd" width="5%">Severity</th>
        <th class="styleTHstd" width="10%">Reason</th>
        <th class="styleTHstd" width="10%">Logged By</th>
    </tr>    

    <%
    If CboDisabled = "" Then 
        For DormantCount = 1 To UBound(DormantArray)        
            Response.Write"<tr>" & VbCrLf
            Response.Write"<td style='color: #0069AA' class='styleTDleft'>" & DormantArray(DormantCount,0) & "</td>" & VbCrLf
            Response.Write"<td  class='styleTDstd'>" & DormantArray(DormantCount,1) & "</td>" & VbCrLf
                    
            If DormantArray(DormantCount,2) = "1" Then
                StatusBg = "#00BB00"
                StatusText = "Minor"
            ElseIf DormantArray(DormantCount,2) = "2" Then
                StatusBg = "#FFA500"  'lighter "#FFBC55" 
                StatusText = "Medium"
            Else
                StatusBg = "#FF0000"
                StatusText = "Critical"
            End If
            
            Response.Write"<td class='styleTDstd' title='" & StatusText & "' style='background-color:" & StatusBg & ";'></td>" & VbCrLf
            Response.Write"<td class='styleTDstd'>" & DormantArray(DormantCount,3) & "</td>" & VbCrLf
            Response.Write"<td class='styleTDstd'>" & DormantArray(DormantCount,4) & "</td>" & VbCrLf	        
            Response.Write"</tr>" & VbCrLf          
	    Next 
	    Erase DormantArray  
    End If 
    %>
</table>
</body>  
</html>