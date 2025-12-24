<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\Declarations.asp" -->
<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->


<% 
'## On Error Resume Next

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")


Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"

Dim DetailRS
Dim strDetailSql
Dim DetailJob
Dim DetailItems
Dim DetailItemsParts
Dim DetailPartSqm
Dim DetailSubPartSqm
Dim DetailOperationStatus
Dim DetailPartDescription
Dim LocalSubstrate
Dim Fabrication(3,1)
Dim FabDummy
Dim ProdNotes
Dim DisplayStatus

Dim OpNotesExist
Dim OperationNotes(6)

Dim PLCWPrivateNotes
Dim DetailPartCode
Dim DetailPartQty
Dim DetailSubPartQty
Dim DetailSubPartItemCount
Dim DetailPrinterName



Dim Client
Dim DeliveryAddress

Dim SubCount
Dim ItemAlpha
Dim Combined
Dim JobIsOutdoor

Dim AlphaItemsParts
Dim AlphaItems

Dim ReqDelHeader
Dim ItemDescription
Dim DetailMaterialCount

Dim MgsColSpan
Dim hQuoteId
Dim hJobId
Dim hJobRef

Dim DeptSql
Dim DeptRs

DeptSql = "SELECT Id, Description, Type, [Order]"
DeptSql = DeptSql & " FROM ProdWorkCentreGroups"
DeptSql = DeptSql & " WHERE (NOT (Id IN (01,08)))"
DeptSql = DeptSql & " ORDER BY [Order]"

Set DeptRs = Server.CreateObject("ADODB.Recordset")

DeptRs.ActiveConnection = strConnClarity
DeptRs.Source = DeptSql
DeptRs.CursorType = 0
DeptRs.CursorLocation = 3
DeptRs.LockType = 1
DeptRs.Open

If DeptRs.BOF = True Or DeptRs.EOF = True Then
    DeptRs.Close
    Set DeptRs = Nothing
    Err.Clear
    Response.Redirect "Error.asp"
End If

Dim DeptArray()
ReDim DeptArray(DeptRs.RecordCount,1)

Dim DeptCount

While Not DeptRs.EOF

    If DeptRs("Id") < 10 Then
        DeptArray(DeptRs.AbsolutePosition,0) = "0" & DeptRs("Id")
    Else
        DeptArray(DeptRs.AbsolutePosition,0) = DeptRs("Id")
    End if
   
    Select Case DeptRs("Id")    
        Case 4
            DeptArray(DeptRs.AbsolutePosition,1) = "Studio"
        Case 6
            DeptArray(DeptRs.AbsolutePosition,1) = DeptRs("Description") & "/Packing"
        Case 7
            DeptArray(DeptRs.AbsolutePosition,1) = DeptRs("Description") & "/Laminating" 
        Case 11
            DeptArray(DeptRs.AbsolutePosition,1) = DeptRs("Description") & "/Testing"
        Case Else
            DeptArray(DeptRs.AbsolutePosition,1) = DeptRs("Description")    
    End Select
  
    DeptRs.MoveNext
Wend

DeptRs.Close
Set DeptRs = Nothing

Dim TextType
Dim InputSize
If Session("ShowText") = Cbool(True) Then
    TextType = "text"
    InputSize = "size = '4'"
Else
    TextType = "hidden"
    InputSize = ""
End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>

    <title>NC Reporting </title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/warning.png" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/NCReportsJSFunc.js"></script>
    <script type="text/javascript" src="JsFiles/NCReportsAjaxFunc.js"></script>
</head>


<body style="padding: 0px; margin: 0px">

<table style="padding-right: 10px; padding-left: 10px; width: 100%;" >
	<tr>
	    
		<td align="left" valign="bottom" height="100" colspan="3">
            <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
        </td> 
	    
    </tr>	
	<tr>
		<td height="8" valign="top" align="left" width="100%" colspan="3">
            <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
        </td>
	</tr>
<!--</table>

<table style="padding-right: 10px; padding-left: 10px; width: 100%;">-->
<tr>
        <td align="left" width="33%" valign="top">
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('NcSelectOption.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to option page</a><br /><br />
        </td> 
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/warning.png" style="width: 20px; height: 20px;" />
            <font style="color: #0069AA; font-weight: bold; font-size: 16px;">New&nbsp;Log</font>
        </td>
        <td align="right" width="33%" valign="top"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>; ">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
    </tr>
</table>

<br />

<form method="post" action="NcUpdateLog.asp"  name="frmAddLog" id="frmAddLog" onsubmit="return NCValidateLogNJ('Add');"> <!--onsubmit="return NCValidateLog('AddNJ');"-->

<table  style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">	

    <tr>
        <td colspan="3" nowrap="nowrap">
            <label style="font-weight: bold; color: <%=NewBlue%>;">Making a selection in a 
            dropdown, enables the next one. The final one enables notes, you can multi 
            select in all</label>
        </td>
    </tr>
    
    <tr>        
        <td colspan="3" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" width="20%">Department</th>
        <th align="left" width="40%" >Selected</th>	        
    </tr>
    
    <tr>
        <td align="left">
            <select id="cboGroupSelect" onchange="javascript:AddGroupNoJob();">
                 <option value="">Select Dept</option>                 
                <%
                    
                    Response.Write "<option value='100'>Sales - CS</option>" & VbCrLf
                    For DeptCount = 1 to Ubound(DeptArray)                    
                        Response.Write "<option value='" & DeptArray(DeptCount,0) & "'>" & DeptArray(DeptCount,1) & "</option>" & VbCrLf
                    Next
                    
                    Response.Write "<option value='101'>Stock</option>" & VbCrLf
                    Response.Write "<option value='0'>Other</option>"
                    
                    Erase DeptArray    
                %>
            </select>
        
        </td>
        <td align="left" colspan="2">        
            <input type="text" id="SelectedGroup" value = ""
                style="border-style: none; width: 98%; display: inline; font-size: 14px; color: #009933"/>
        </td>        
    </tr> 
    
    <tr>        
        <td colspan="3" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" width="20%">Issue/Problem</th>
        <th align="left" width="40%">Selected</th>
        <th align="left" width="40%" >&nbsp;</th>	        
    </tr>
    
    <tr>
        <td align="left">
            <select id="cboReason" onchange="javascript:AddReason();" disabled="disabled">
                 <option value="">Select Issue/Problem</option>
                           
            </select>        
        </td>
        <td align="left" colspan="2">
            <input type="text" id="SelectedReason" value = ""
                style="border-style: none; width: 98%; display: inline; font-size: 14px; color: #009933"/>        
        </td>         
    </tr>        
    
    <tr>        
        <td colspan="3" align="left">&nbsp;</td>
    </tr>    
    
    
    <tr>       
    <th  colspan="3" align="left">Notes</th>
    </tr>        
    
    <tr>
        <td colspan="3">
            <textarea  rows="15" id="txtDetails" name="txtDetails" disabled="disabled" style="text-align: left; " cols="110" onkeyup="javascript:EnableSubmit();"></textarea>
        </td>
    </tr>
  
</table>

<br />

<table style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">
    <tr>
        
        <td>
            <input type="submit" value="Update" disabled="disabled" id="btnSubmit"/>        
            &nbsp;&nbsp;&nbsp;
            <input name="btnReset" id="btnReset" type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            
            <%
            If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                If TextType = "text" Then
                    Response.Write ("&nbsp;Don't Save Data&nbsp;&nbsp;")
                    Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                    Response.Write ("<br />")
                End If                    
            End If          
            %>   
            <br /><br /> 
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hGroup"%> 
            <input type="<%=TextType%>" id="hGroup" name="hGroup" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hGroupSelected"%> 
            <input type="<%=TextType%>" id="hGroupSelected" name="hGroupSelected" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hReason"%> 
            <input type="<%=TextType%>" id="hReason" name="hReason" value=""  />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hReasonSelected"%> 
            <input type="<%=TextType%>" id="hReasonSelected" name="hReasonSelected" value=""  />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hUserId"%> 
            <input type="<%=TextType%>" id="hUserId" name="hUserId" value="<%=Session("UserId")%>" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hJobType"%> 
            <input type="<%=TextType%>" id="hJobType" name="hJobType" value="0" <%=InputSize %> />
            &nbsp;&nbsp;<input type="<%=TextType%>" id="frmName" name="frmName" value="frmNcAddNewNJ" />
            
           
        </td>
    </tr>
    <tr>
        <td id="PrivateNotes" style="visibility: visible">&nbsp;</td>
    </tr>    
</table>

</form>

</body>

</html>

