<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="HsGetData.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->


<% 
'## On Error Resume Next

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect "SystemLocked.asp"

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim LogType
LogType = Cint(Request.QueryString("st"))

Dim ResolvedFilter
ResolvedFilter = Cint(Request.QueryString("rf"))


Dim OkStatus 
Dim ReturnAddress
Dim Digital
Dim Titletext
OkStatus = 0
Dim EditId
EditId = ""
Dim NoDisplay
Dim FlushCount
Dim Msg
Dim StatusBg
Dim StatusText
Dim HeaderText
Dim LoggedByName
Dim RiddorTitle

Dim NotesTitleText

If LogType = 1 Then
    HeaderText = "All Accident Records"
ElseIf LogType = 2 Then
    HeaderText = "All Near Miss Records"
Else
    HeaderText = "All Unsafe Condition or Damage Records"
End If

If ResolvedFilter = 0 Then HeaderText = Replace(HeaderText,"All ", "Unresolved ",1,-1,1)

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
    LockRs.Source = LockRs.Source & " And (SystemName = 'HS')"
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
    Session("LockedBy") = ""
    Session("LockedByName") = ""
    
End If

Session("AddFolder") = ""


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
            &nbsp;&nbsp;<a href ="javascript:GoBackHSOption();" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;select&nbsp;option&nbsp;page</a>            
        </td>
        <td align="center" valign="bottom" width="34%">
            <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >&nbsp;<%=HeaderText%></font>                     
        </td>
        <td align="right" width="33%">
            <a href ="javascript:LogOff()" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;Out</a>&nbsp;&nbsp;    
        </td>
    </tr>
   
    <tr>     
        <td align="left" valign="middle" colspan="3">        
            <%
            If Session("AdminUser") = Cbool(True) Or Session("EditHS") = Cbool(True) Then
                Response.Write "<br />&nbsp;To&nbsp;edit&nbsp;a&nbsp;record&nbsp;click&nbsp;anywhere&nbsp;in&nbsp;the&nbsp;Record&nbsp;Id&nbsp;area"     
            Else 
                Response.Write "&nbsp;"                           
            End If
            %>   
        </td>
    </tr>
</table>


<%

Call OpenReasons(LogType)

While Not ReasonRs.Eof 
    Msg = "<h4 style = 'color: Black' >&nbsp;&nbsp;" & ReasonRs("Description") & " </h4>"   
      
    Call OpenLogs (LogType, ReasonRs("Id"), ResolvedFilter)    
    '## If no records don't draw tables
    If LogRs.Bof = False or LogRs.EOF = False Then
        Response.Write (Msg)
    %>    
    <table  align="center" cellpadding="0" cellspacing="0"  width="95%">    
	   <tr >
	        <th class="styleTHleft" width="5%" >Record Id</th>
	        <th class="styleTHstd" width="10%">Created</th>	        
	        <th class="styleTHstd" width="5%">Severity</th>
	        <th class="styleTHstd" width="5%">Riddor</th>
	        <th class="styleTHstd" width="10%">Logged By</th>
	        <th class="styleTHstd" width="4%">Images</th>
	        <th class="styleTHstd" width="15%">Reason</th>
	        <th class="styleTHstd" width="5%">Resolved</th>
	        <th class="styleTHstd" width="10%">Date/Time of Incident</th>
	        <th class="styleTHstd" width="10%">Last Updated</th>
        </tr>    
                
        <%
        While Not LogRs.Eof
        
        If Session("AdminUser") = Cbool(True) Or Session("EditHS") = Cbool(True) Then
            EditId = LogRs("Id")
        Else
            EditId = ""
        End If
        
        NotesTitleText = Replace(LogRs("Notes"),"<br />",VbCrLf,1,-1,1) 
        %>        
        <tr>
            <td title="" id="<%=EditId%>" style="color: #0069AA" class="styleTDleft"
                <%If Session("AdminUser") = Cbool(True) Or Session("EditHS") = Cbool(True) Then Response.Write " onclick='javascript:HSLoadEditOrder(" & EditId & ");'" %> ><%=LogRs("Id")%></td>
	        <td  class="styleTDstd"><%=Cdate(LogRs("CreatedDate"))%></td>	    
	        <%
	            If LogRs("Severity") = 1 Then
	                StatusBg = "#00BB00"
	                StatusText = "Minor"
	            ElseIf LogRs("Severity") = 2 Then
	                StatusBg = "#FFA500"  'lighter "#FFBC55" 
	                StatusText = "Medium"
	            Else
	                StatusBg = "#FF0000"
	                StatusText = "Critical"
	            End If
	            
	            LoggedByName = Trim(LogRs("CreatedByName"))
	        %>
	        	        
	        
	        <td title="<%=StatusText%>" class="styleTDstd" style="background-color:<%=StatusBg%>;"></td>
	        
	        <td  class="styleTDstd">
	        <%
	        If LogRs("Riddor") = Cbool(True) Then
	            If LogRs("RiddorSubmitted") = Cbool(True) Then
	                RiddorTitle = "Riddor submitted on " & Cdate(LogRs("RiddorSubmitedDate"))            
	                Response.Write "<img align='middle' alt='' title='" & RiddorTitle & "' src='Images/checkmark.ico' width='15' height='15'  />"
	            Else
	                Response.Write "<img align='middle' alt='' title='Pending Submission' src='Images/X2.ico' width='15' height='15'  />"
	            End If
	        Else
	            Response.Write "&nbsp;"
	        End If	        
	        %>	        
	        </td>  
	        
	        
	        
	        <td  class="styleTDstd"><%=LoggedByName%></td>
	        
	        <td  class="styleTDstd">       
	        <%
	        If LogRs("HasImage") = Cbool(True) Then
	        	Response.Write "<img align='middle' alt='' title='Click to view images' src='Images/checkmark.ico' width='15' height='auto' onclick='javascript:LoadImage(" & EditId & ");' />"
   	        	'Response.Write "<img align='middle' alt='' title='Click to view images' src='Images/camera.ico' width='20' height='auto' onclick='javascript:LoadImage(" & EditId & ");' />"
	        Else
	            Response.Write "&nbsp;"
	        End If	        
	        %>	        
	        </td>
	        
	        <td class="styleTDstd" title="<%=NotesTitleText%>" ><%=LogRs("Description")%></td>
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
        LoggedByName = ""
        EditId = ""
        NotesTitleText = ""
        LogRs.MoveNext 
        
        FlushCount = FlushCount + 1
        If FlushCount >= 50 Then
            Response.Flush
            FlushCount = 0
        End If
                                             
        Wend
        %>
        
        </table>
        <br />
        <hr style="border-style: none; width: 98%; height: 1px; background-color: <%=NewCyan%>; display: block;" />         
        
        <%
        LogRs.Close
        Set LogRs = Nothing
                   
      End If

        ReasonRs.MoveNext 
    Wend

    ReasonRs.Close  
    Set ReasonRs = Nothing   
        
%>   
<br /><br /><br />
<table>
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
If Session("Locked") = Cbool(True) Then 
    If Application("UsersOnLine") > 0 Then
        Application("UsersOnLine") = Application("UsersOnLine") -1
    End If
    Session.Abandon 
End If
%>