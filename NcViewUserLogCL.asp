<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\Declarations.asp" -->
<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="NcGetData.asp" -->
<!--#include file="CommonFunctions.asp" -->

<% 
'## On Error Resume Next

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")


Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"

Dim LogToGet
LogToGet = Clng(Request.QueryString("id"))


Call OpenLogs(LogToGet)

Dim TextType
Dim InputSize
If Session("ShowText") = Cbool(True) Then
    TextType = "text"
    InputSize = "size = '4'"
Else
    TextType = "hidden"
    InputSize = ""
End If



Dim LocalNotes
LocalNotes = Trim(LogRS("Notes"))

If LocalNotes <> "" Then
    LocalNotes = Replace(LocalNotes,"<br />",VbCrLf,1,-1,1)
End If

Dim LocalResponseNotes
LocalResponseNotes = Trim(LogRS("ResponseNotes"))

If LocalResponseNotes <> "" Then
    LocalResponseNotes = Replace(LocalResponseNotes,"<br />",VbCrLf,1,-1,1)
Else
    LocalResponseNotes = ""
End If


Dim LoggedByName
LoggedByName = Trim(LogRs("CreatedByName"))

Dim ResolvedByName
ResolvedByName = GetMcUser (LogRs("ResolvedBy"))

Dim LastModifiedByName
LastModifiedByName = GetMcUser (LogRs("LastModifiedBy"))

Dim ResponseRows
Dim ResponseMsg
Dim ResponseHeaderExtra
Dim LockResponseText
Dim NewResponseHtml
Dim LocalHasNotes

If LogRs("LastModifiedDateSerial") > 0 Then
    ResponseMsg = "Last update by " & LastModifiedByName & "&nbsp;" & Cdate(LogRs("LastModifiedDate"))
    If LocalResponseNotes = "" Then
        ResponseMsg = ""
        LockResponseText = ""
        NewResponseHtml = ""
        ResponseRows = 15
        ResponseHeaderExtra = "*"
        LocalHasNotes = False
    Else
        LocalHasNotes = True
        LockResponseText = " readonly='readonly' "
        ResponseRows = 8
        ResponseHeaderExtra = ""
        NewResponseHtml = "<label style=' color: " & NewBlue & ";'>New&nbsp;Comments</label><br />"
        NewResponseHtml = NewResponseHtml & " <textarea  onkeyup='javascript:EnableSubmit()' rows='8' id='txtResponseNew' "
        NewResponseHtml = NewResponseHtml & " name='txtResponseNew' style='text-align: left; ' cols='70' ></textarea>"
    End If
Else
    ResponseMsg = ""
    LockResponseText = ""
    NewResponseHtml = ""
    ResponseRows = 15
    ResponseHeaderExtra = "*"
    LocalHasNotes = False
End If

Dim MadeUpWorkCentre
Dim ClarityWorkcentre
ClarityWorkcentre = DeptDescription (LogRs("DeptIds"))

MadeUpWorkCentre = LogRs("DeptIds")

If ClarityWorkcentre = "" Then
    If Instr(1,MadeUpWorkCentre,"100") > 0 Then ClarityWorkcentre = "Sales - CS"    
    If Instr(1,MadeUpWorkCentre,"101") > 0 Then ClarityWorkcentre = ClarityWorkcentre & ", Stock" 
    If Instr(1,MadeUpWorkCentre,"102,") > 0 Then ClarityWorkcentre = ClarityWorkcentre & ", Other"
    If LogRs("DeptIds") = "0" Then
        ClarityWorkcentre = "Other"
    Else
        If Left(LogRs("DeptIds"),1) = "0" Then ClarityWorkcentre = ClarityWorkcentre & ", Other"
    End If
Else
    If Instr(1,MadeUpWorkCentre,"100") > 0 Then ClarityWorkcentre = "Sales - CS, " & ClarityWorkcentre
    If Instr(1,MadeUpWorkCentre,"101") > 0 Then ClarityWorkcentre = ClarityWorkcentre & ", Stock"
    If Instr(1,MadeUpWorkCentre,"102,") > 0 Then ClarityWorkcentre = ClarityWorkcentre & ", Other"
    If LogRs("DeptIds") = "0" Then
        ClarityWorkcentre = "Other"
    Else
        If Left(LogRs("DeptIds"),1) = "0" Then ClarityWorkcentre = ClarityWorkcentre & ", Other"
    End If
End If 

ClarityWorkcentre = Replace(ClarityWorkcentre,",,",",",1,-1,1) 


Dim LocalItemsAlpha
LocalItemsAlpha = LogRs("QuoteItemsAlpha")  
LocalItemsAlpha = Replace(LocalItemsAlpha,",", ", ",1,-1,1)

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
</head>


<body style="padding: 0px; margin: 0px" >

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

    <tr>
        <td align="left" width="33%" valign="top">
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('NcSelectOption.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to option page</a>
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('NcDisplayUserLogs.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to view logs</a><br /><br />
        </td> 
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/warning.png" style="width: 20px; height: 20px;" />
            <font style="color: #0069AA; font-weight: bold; font-size: 16px;">
                View&nbsp;Log&nbsp#<%=LogToGet%>&nbsp;-&nbsp;<%=LogRs("QuoteRef") %>
            </font>
        </td>
        <td align="right" width="33%" valign="top"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>; ">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
    </tr>
</table>

<br />

<form method="post" action="NcUpdateLog.asp"  name="frmEditLogCL" id="frmEditLogCL" onsubmit="return NCValidateLogCL('Edit');"> 

<table  style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">	

    <tr >		 
        <th align="left" colspan="4">Items </th>                
    </tr>
    
    <tr>
        <td align="left" colspan="4">        
            <input type="text" id="SelectedItem" value="<%=LocalItemsAlpha%>"
                style="border-style: none; width: 95%; display: inline; font-size: 14px; color: #009933"/>
        </td>
    </tr>
        
    <tr>        
        <td colspan="4" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" colspan="4">Work Centre</th>        	        
    </tr>
    
    <tr>
        <td align="left" colspan="4">        
            <input type="text" id="SelectedDept" value = "<%=ClarityWorkcentre%>"
                style="border-style: none; width: 95%; display: inline; font-size: 14px; color: #009933"/>
        </td>
    </tr> 
    
    <tr>        
        <td colspan="4" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" colspan="4">Issue/Problem</th>                
    </tr>
    
    <tr>
        <td align="left" colspan="4">
            <input type="text" id="SelectedReason" value = "<%=ReasonDescription (LogRS("ReasonSelection")) %>"
                style="border-style: none; width: 95%; display: inline; font-size: 14px; color: #009933"/>        
        </td>         
    </tr>        
            
    <tr>        
        <td colspan="4" align="left">&nbsp;</td>
    </tr>
    
    <tr>
        <td align="left" width="50%" colspan="2">
            <label style=" color: <%=NewBlue%>; font-size: medium;">Created by <%=LoggedByName%>&nbsp;<%=LogRs("SelectedDate") %></label>
            <br /> <br />    
        </td>
        <td align="left" width="50%" colspan="2">
            <label style=" color: <%=NewBlue%>; font-size: medium;"><%=ResponseMsg%></label> 
            <br /> <br />    
        </td>
    </tr>
              
    <tr>
        <td colspan="2">
        <label style=" color: <%=NewBlue%>">Details</label><br />
            <textarea  rows="20" id="txtDetails" name="txtDetails" readonly="readonly" style="text-align: left; " cols="60" ><%=Trim(LocalNotes)%></textarea>
        </td>
        <td colspan="2">            
            <label style=" color: <%=NewBlue%>">Comments<font style="color: red"><%=ResponseHeaderExtra%></font></label><br />
            <textarea  rows="20" id="txtResponse" readonly="readonly"
                name="txtResponse" style="text-align: left; " cols="70" ><%=Trim(LocalResponseNotes)%></textarea>
            <br />
                  
        </td>
    </tr>
</table>

<br />

<table style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">
    <tr>        
        <td>
            
            <label style="font-size: 14px; font-weight: bold;" >&nbsp;Resolved</label>
            <input type="checkbox" id="chkResolved" name="chkResolved" disabled="disabled"
                <%
                If LogRs("Resolved") = CBool(True) Then
                    Response.Write " checked='checked' "                
                    Response.Write " title='Marked resolved on " & Cdate(LogRs("ResolvedDateSerial")) & " by " & ResolvedByName & "' "
                End If
                %> 
             />
            

            <br /><br /> 
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hLogId"%>
            <input type="<%=TextType%>" id="hLogId" name="hLogId" value="<%=LogToGet%>" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hUserId"%> 
            <input type="<%=TextType%>" id="hUserId" name="hUserId" value="<%=Session("UserId")%>" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hJobType"%> 
            <input type="<%=TextType%>" id="hJobType" name="hJobType" value="1" <%=InputSize %> />
            &nbsp;&nbsp;<input type="<%=TextType%>" id="frmName" name="frmName" value="frmNcViewLogCL" />
            
            <%If TextType = "text" Then Response.Write "hHasNotes"%>
            <input type="<%=TextType%>" name="hHasNotes" id="hHasNotes" value="<%=LocalHasNotes%>" <%=InputSize %>/>
            
           
        </td>
    </tr>
    <tr>
        <td id="PrivateNotes" style="visibility: visible">&nbsp;</td>
    </tr>    
</table>

</form>

</body>

</html>

