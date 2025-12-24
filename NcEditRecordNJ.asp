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

'## Check If Locked
Dim LockStatus
LockStatus = ChkRecordLock

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

<body style="padding: 0px; margin: 0px" onload="javascript:NCEditLogLoad();">

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
    <tr>
        <td align="left" width="33%" valign="top">
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('NcSelectOption.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to option page</a>
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('NcDisplay.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to view logs</a><br /><br />
        </td> 
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/warning.png" style="width: 20px; height: 20px;" />
            <font style="color: #0069AA; font-weight: bold; font-size: 16px;">Edit&nbsp;Log&nbsp#<%=LogToGet%></font>
        </td>
        <td align="right" width="33%" valign="top"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>; ">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
    </tr>
</table>

<br />

<form method="post" action="NcUpdateLog.asp"  name="frmEditLogNJ" id="frmEditLogNJ" onsubmit="return NCValidateLogNJ('Edit');"> 

<table  style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">        
    <tr>        
        <td colspan="4" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" width="20%">Department</th>
        <th align="left" width="40%" colspan="2">&nbsp;</th>
        <th align="left" width="40%" >&nbsp;</th>	        
    </tr>
    
    <tr>
        <td align="left" colspan="4">        
            <input type="text" id="SelectedGroup" value = "<%=GroupDescription (LogRs("GroupSelection"))%>"
                style="border-style: none; width: 95%; display: inline; font-size: 14px; color: #009933"/>
        </td>
    </tr> 
    
    <tr>        
        <td colspan="4" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" width="20%">Issue/Problem</th>
        <th align="left" width="40%" colspan="2">&nbsp</th>
        <th align="left" width="40%" >&nbsp;</th>	        
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
        <textarea  rows="15" id="txtDetails" name="txtDetails" readonly="readonly" style="text-align: left; " cols="60" ><%=Trim(LocalNotes)%></textarea>
   
        </td>
        <td colspan="2">            
            <label style=" color: <%=NewBlue%>">Comments<font style="color: red"><%=ResponseHeaderExtra%></font></label><br />
            <textarea  onkeyup="javascript:EnableSubmit()" rows="<%=ResponseRows%>" id="txtResponse" <%=LockResponseText%>
                name="txtResponse" style="text-align: left; " cols="70" ><%=Trim(LocalResponseNotes)%></textarea>
            <br />
            <%=NewResponseHtml%>       
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
            &nbsp;&nbsp;&nbsp;
            <label style="font-size: 14px; font-weight: bold;" >&nbsp;Resolved</label>
            <input type="checkbox" id="Checkbox1" name="chkResolved" onclick="javascript:EnableSubmit()"
                <%If LogRs("Resolved") = CBool(True) Then Response.Write " checked='checked' "%> 
                <%If LogRs("Resolved") = CBool(True) Then Response.Write " disabled='disabled' "%>
                <%
                If LogRs("Resolved") = CBool(True) Then 
                    Response.Write " title='Marked resolved on " & Cdate(LogRs("ResolvedDateSerial")) & " by " & ResolvedByName & "' "
                End If
                %> 
             />
            
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
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hLogId"%>
            <input type="<%=TextType%>" id="hLogId" name="hLogId" value="<%=LogToGet%>" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hUserId"%> 
            <input type="<%=TextType%>" id="hUserId" name="hUserId" value="<%=Session("UserId")%>" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hJobType"%> 
            <input type="<%=TextType%>" id="hJobType" name="hJobType" value="0" <%=InputSize %> />
            &nbsp;&nbsp;<input type="<%=TextType%>" id="frmName" name="frmName" value="frmNcEditLogNJ" />
            <%If TextType = "text" Then Response.Write "hLockStatus"%>
            <input type="<%=TextType%>" name="hLockStatus" id="hLockStatus" value= "<%=LockStatus%>"  <%=InputSize %>/>&nbsp;
            <%If TextType = "text" Then Response.Write "LockedByName"%> 
            <input type="<%=TextType%>" name="LockedByName" id="LockedByName" value= "<%=Session("LockedByName")%>" <%=InputSize %>/>&nbsp;
            <%If TextType = "text" Then Response.Write "Locked"%> 
            <input type="<%=TextType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" <%=InputSize %>/>&nbsp;
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

<%

LogRS.Close
Set LogRS = Nothing

Private Function ChkRecordLock()

    '## Set locked status, so that other users can't open 

    Dim RecordLockSql
    Dim RecordLockRs
    Dim LocalRecordLock
    LocalRecordLock = False
    
    Set RecordLockRs = Server.CreateObject("ADODB.Recordset")
    RecordLockRs.ActiveConnection = Session("ConnMcLogon")
    RecordLockRs.Source = "Select * From RecordLocks"  
    RecordLockRs.Source = RecordLockRs.Source & " Where (RecordId = " & LogToGet & ") AND (SystemName = 'NC')"
    RecordLockRs.CursorType = Application("adOpenStatic")
    RecordLockRs.CursorLocation = Application("adUseClient")
    RecordLockRs.LockType = Application("adLockOptimistic")
    RecordLockRs.Open
    
    If RecordLockRs.BOF = True Or RecordLockRs.EOF = True Then
        '## No match add new
        RecordLockRs.AddNew
        RecordLockRs("RecordId") = LogToGet
        RecordLockRs("LockedbyId") = Session("UserId")
        RecordLockRs("LockedByName") = Session("UserName")
        RecordLockRs("SystemName") = "NC"
        Session("LockedByName") = Session("UserName")
        Session("LockedId") = LogToGet
        
        RecordLockRs.Update
        RecordLockRs.MoveLast        
    Else
        '## Already locked, do nothing
        If RecordLockRs("LockedById") = Session("UserId")  Then  'And RecordLockRs("SystemName") = "HS"
            '## Match found but was just a reset of page by same user
            LocalRecordLock = False
        Else
            '## Match found so status is true
            LocalRecordLock = True
            Session("LockedByName") = RecordLockRs("LockedByName")  
        End If
    End If
    
    Set RecordLockRs.ActiveConnection = Nothing
    RecordLockRs.Close
    Set RecordLockRs = Nothing
    
    ChkRecordLock = LocalRecordLock

End Function



%>