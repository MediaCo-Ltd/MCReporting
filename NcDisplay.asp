<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="NcGetData.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<% 
'## On Error Resume Next

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")


Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim LogType

LogType = Cint(Request.QueryString("st"))

'## If Cookie already set don't bother with check

Dim OkStatus 
Dim ReturnAddress
Dim Digital
Dim NotesTitleText

OkStatus = 0
Dim EditId
EditId = ""
Dim JobType
JobType = ""
Dim NoDisplay
Dim FlushCount
Dim Msg
Dim StatusBg
Dim StatusText
Dim HeaderText

HeaderText = "All Records"

Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"    
End If


If Session("LockedId") <> "" Then
    '## Clear record lock, but only if made by this user
       
    Dim LockSql
    Dim LockRs

    Set LockRs = Server.CreateObject("ADODB.Recordset")
    LockRs.ActiveConnection = Session("ConnMcLogon")
    LockRs.Source = "Select * From RecordLocks"          
    LockRs.Source = LockRs.Source & " Where (RecordId = " & Session("LockedId") & ")"
    LockRs.Source = LockRs.Source & " And (LockedById = " & Session("UserId") & ")"
    LockRs.Source = LockRs.Source & " And (SystemName = 'NC')"
    LockRs.CursorType = Application("adOpenStatic")
    LockRs.CursorLocation = Application("adUseClient")
    LockRs.LockType = Application("adLockOptimistic")
    LockRs.Open

    If LockRs.BOF = True Or LockRs.EOF = True Then
        '## No match so do nothing
    Else
        If LockRs("LockedbyId") = Session("UserId") Then
            LockRs.Delete
            LockRs.MoveNext
        End If   
    End If

    Set LockRs.ActiveConnection = Nothing
    LockRs.Close
    Set LockRs = Nothing

    Session("LockedId") = ""
    Session("LockedSid") = ""
    Session("LockedBy") = ""
    Session("LockedByName") = ""
End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>NC Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/NCReportsJSFunc.js"></script>
    <link rel="shortcut icon" href="Images/warning.png" type="image/x-icon" /> 
</head>

<body style="padding: 0px; margin: 0px" id="Display" >  <!--onload="javascript:ViewRecordsLoad();"-->

<table style="width: 100%;padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100" colspan="3" >
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />
        </td>
    </tr>
    <tr>
	    <td height="8" valign="top" colspan="3" >
	        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
	    </td>
    </tr>    
	
    <tr>
        <td align="left" style="height: 20px" valign="bottom" width="33%" >
            &nbsp;&nbsp;<a href ="javascript:GoBackNCOption();" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;select&nbsp;option&nbsp;page</a>            
        </td>
        <td align="center" valign="bottom" width="34%">
            <img align="top" alt="mediaco logo" src="Images/warning.png" style="width: 20px; height: 20px;" /> 
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >&nbsp;<%=HeaderText%></font>                     
        </td>
        <td align="right" width="33%">
            <a href ="javascript:LogOff()" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;Out</a>&nbsp;&nbsp;    
        </td>
    </tr>
   
    <tr>     
        <td align="left" valign="middle" colspan="3">        
            <%
            If Session("AdminUser") = Cbool(True) Or Session("EditNC") = Cbool(True) Then
                Response.Write "<br />&nbsp;To&nbsp;edit&nbsp;a&nbsp;record&nbsp;click&nbsp;anywhere&nbsp;in&nbsp;the&nbsp;Record&nbsp;Id&nbsp;area"
            Else 
                Response.Write "&nbsp;"                
            End If
            %>   
        </td>
    </tr>
</table>


<%



'While Not ReasonRs.Eof 
    Msg = "<h4 style = 'color: Black' >&nbsp;&nbsp;All records</h4>"   
      
    Call OpenLogs ("") '## (LogType, ReasonRs("Id"))    
    '## If no records don't draw tables
    
        Response.Write (Msg)
    %>    
    <table  align="center" cellpadding="0" cellspacing="0"  width="95%">    
	   <tr >
	        <th class="styleTHleft" width="4%" >Record Id</th>
	        <th class="styleTHstd" width="5%">Created</th>	        
	        <th class="styleTHstd" width="5%">Job Ref</th>
	        <th class="styleTHstd" width="8%">Logged By</th>
	        <th class="styleTHstd" width="25%">Department</th>
	        <th class="styleTHstd" width="5%">Resolved</th>
	        <th class="styleTHstd" width="7%">Date/Time of Incident</th>
	        <th class="styleTHstd" width="7%">Last Updated</th>
        </tr>    
    <!--</table>
    
    <table  align="center" cellpadding="0" cellspacing="0"  width="95%">-->           
        <%
        
   If LogRs.Bof = False or LogRs.EOF = False Then
        
        While Not LogRs.Eof
        
        If Session("AdminUser") = Cbool(True) Or Session("EditNC") = Cbool(True) Then
            EditId = LogRs("Id")
            JobType = LogRs("JobTypeId")
        Else
            EditId = ""
            JobType = ""
        End If 
        
        NotesTitleText = Replace(LogRs("Notes"),"<br />",VbCrLf,1,-1,1)
        %> 
        
             
        <tr>
            <td title="" id="<%=EditId%>" style="color: #0069AA" class="styleTDleft"
                <%If Session("AdminUser") = Cbool(True) Or Session("EditNC") = Cbool(True) Then
                     Response.Write " onclick=' javascript:NcEditRecord(" & EditId & "," & JobType & ");' "
                       
                  End If
                 %>   
                     ><%=LogRs("Id")%>
            </td>
                
	        <td  class="styleTDstd"><%=Cdate(LogRs("CreatedDate"))%></td>	    
	        	        
	        <td  class="styleTDstd" ><%=LogRs("QuoteRef") %></td>
	        <td  class="styleTDstd"><%=LogRs("CreatedByName")%></td>
	        <td class="styleTDstd" title="<%=NotesTitleText%>" style="text-align: left; padding-left: 5px;">
	            <%Response.Write GroupDescription (LogRs("GroupSelection"))%>
	        </td>
	        <td  class="styleTDstd">
	        <%
	        If LogRs("Resolved") = Cbool(True) Then
	            Response.Write "<img align='middle' alt='' title='Resolved' src='Images/checkmark.ico' width='15' height='15'  />"
	        Else
	            Response.Write "<img align='middle' alt='' title='Pending' src='Images/X2.ico' width='15' height='15'  />"
	        End If	        
	        %>	        
	        </td>
	        <td  class="styleTDstd"><%=LogRs("SelectedDate")%></td>
	        <td  class="styleTDstd">
	            <%If LogRs("LastModifiedDateSerial") > 0 Then Response.Write Cdate(LogRs("LastModifiedDateSerial"))%>
	        </td>
	      </tr>
	      
        <%
        
        OkStatus = 0
        StatusBg = ""
        StatusText = ""
        NotesTitleText = ""
       
        EditId = ""
        LogRs.MoveNext 
        
       ' FlushCount = FlushCount + 1
       ' If FlushCount >= 50 Then
        '    Response.Flush
        '    FlushCount = 0
        'End If
                                             
        Wend
        %>
        
        
   
        <!--<br />
        <hr style="border-style: none; width: 98%; height: 1px; background-color: <%=NewCyan%>; display: block;" />  -->       
        
        <%
        
                   
 End If
      
      LogRs.Close
        Set LogRs = Nothing

    '    ReasonRs.MoveNext 
    'Wend

   ' ReasonRs.Close  
   ' Set ReasonRs = Nothing   
        
%>
</table>

<br /><br /><br />

   


<table  align="center" cellpadding="0" cellspacing="0"  width="95%">  
<tr>
<td>
<%If InputType = "text" Then Response.Write "Locked"%> 
<input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" />&nbsp;
<%If InputType = "text" Then Response.Write "Sendpage"%> 
<input type="<%=InputType %>" name="Sendpage" id ="Sendpage" value="<%=LogType%>" />&nbsp;
</td>
</tr>
</table>
</body>
</html>
<% 
'If Session("Locked") = Cbool(True) Then 
'    If Application("UsersOnLine") > 0 Then
'        Application("UsersOnLine") = Application("UsersOnLine") -1
'    End If
'    Session.Abandon 
'End If



%>