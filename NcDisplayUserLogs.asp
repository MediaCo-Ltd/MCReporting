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

'LogType = Cint(Request.QueryString("st"))

'## If Cookie already set don't bother with check

Dim OkStatus 
Dim ReturnAddress
Dim Digital
Dim NotesTitleText
OkStatus = 0
Dim EditId
EditId = ""
Dim NoDisplay
Dim FlushCount
Dim Msg
Dim StatusBg
Dim StatusText
Dim HeaderText

HeaderText = "All Records created by " & Session("UserName")

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
            <br />&nbsp;To&nbsp;view&nbsp;a&nbsp;record&nbsp;click&nbsp;anywhere&nbsp;in&nbsp;the&nbsp;Record&nbsp;Id&nbsp;area
        </td>
    </tr>
</table>


<%



'While Not ReasonRs.Eof 
    Msg = "<h4 style = 'color: Black' >&nbsp;&nbsp;All Records</h4>"   
      
    Call OpenLogByUser  
    '## If no records don't draw tables
    
        Response.Write (Msg)
    %>    
    <table  align="center" cellpadding="0" cellspacing="0"  width="95%">    
	   <tr >
	        <th class="styleTHleft" width="4%" >Record Id</th>
	        <th class="styleTHstd" width="5%">Created</th>	        
	        <th class="styleTHstd" width="5%">Job Ref</th>
	        <th class="styleTHstd" width="8%">Logged By</th>
	        <th class="styleTHstd" width="40%">Department</th>
	        <th class="styleTHstd" width="5%">Resolved</th>
	        <th class="styleTHstd" width="7%">Date/Time of Incident</th>
	        <th class="styleTHstd" width="7%">Last Updated</th>
        </tr>    
              
        <%
        
   If LogRs.Bof = False or LogRs.EOF = False Then
        
        While Not LogRs.Eof
            EditId = LogRs("Id") & "," & LogRs("JobTypeId")
            NotesTitleText = Replace(LogRs("Notes"),"<br />",VbCrLf,1,-1,1)
        %> 
        
             
        <tr>
            <td title="" style="color: #0069AA" class="styleTDleft"
                onclick="javascript:NcViewRecordUser(<%=EditId%>);"><%=LogRs("Id")%>
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
       
        EditId = ""
        LogRs.MoveNext 
       
                                             
        Wend
              
                   
 End If
      
      LogRs.Close
      Set LogRs = Nothing
 
        
%>
</table>

<br /><br /><br />

   


<table  align="center" cellpadding="0" cellspacing="0"  width="95%">  
<tr>
<td>
<%If InputType = "text" Then Response.Write "Locked"%> 
<input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" />&nbsp;

</td>
</tr>
</table>
</body>
</html>
