<%@language="vbscript" codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->



<% 
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"
If Session("ConnHSReports") = "" Then Response.Redirect("Admin.asp")

Dim ShowIp
Dim EditImg
Dim UserRs
Dim UserSql
Dim RowCount
Dim UserCount(2)
UserCount(0) = 0
UserCount(1) = 0
UserCount(2) = 0

Dim LogCount(3)
LogCount(0) = 0
LogCount(1) = 0
LogCount(2) = 0
LogCount(3) = 0


'## 127.0.0.1 Local Pc
'## 192.168.20.110 My Pc at work to server
                                                                                                                                                               	
If Request.ServerVariables("REMOTE_ADDR") = "127.0.0.1" Or Request.ServerVariables("REMOTE_ADDR") = "192.168.20.110" Then
    ShowIp = True
Else
    ShowIp = False
End If
                    '## use this to hide me Where (NOT (Id IN (1, 4)))    
UserSql = "Select Id, UserName, Password, AdminUser, Active, ShowHS, ShowNC, ShowMF, HsEdit, HsView, NcEdit, MfEdit, RedoAdd, RedoView, Dormant from Users Where((NOT (Id = 1))) AND (Dormant = 1) Order by UserName"   
    
Set UserRs = Server.CreateObject("ADODB.Recordset")

UserRs.ActiveConnection = Session("ConnMcLogon")
UserRs.Source = UserSql
UserRs.CursorType = Application("adOpenForwardOnly")        
UserRs.CursorLocation = Application("adUseClient")   
UserRs.LockType = Application("adLockReadOnly")         
UserRs.Open 
Set UserRs.ActiveConnection = Nothing

Dim DormantToolTip

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/PenPad64.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/MCReportsJSFunc.js"></script>   
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
            &nbsp;<a href ="javascript:window.location.replace('Admin.asp');" style="color:<%=NewBlue%>">Return&nbsp;to&nbsp;admin&nbsp;page</a>
        </td>
	    <td  align="center" width="34%">
	        <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >MediaCo&nbsp;Reporting&nbsp;Dormant&nbsp;Users</font>
	    </td>
	    <td align="right" width="33%">&nbsp;</td>
	</tr>    
</table>


<table style="width: 100%; position: absolute; top: 18%; padding-right: 20px; padding-left: 20px;">
    <tr>
        <td width="10%">&nbsp;</td>
        <td width="15%">
            &nbsp;
        </td>
        <td width="60%">&nbsp;<font size="2" style="color: #0069AA">Click&nbsp;Name&nbsp;to&nbsp;edit&nbsp;user.</font></td>                 
    </tr>
</table>

<table style="width: 90%; position: absolute; top: 25%; left: 5%;" cellpadding="0" cellspacing="0" >  <!--left: 5%;-->
    <tr>
        <th class="styleTHleft" width="11%">User</th>
        <th class="styleTHstd" width="10%">Password</th>
        <th class="styleTHstd" width="5%">Id</th>
        <th class="styleTHstd" width="5%">Status</th>        
        <th class="styleTHstd" width="5%">HS User</th>
        <th class="styleTHstd" width="5%">HS View</th>
        <th class="styleTHstd" width="5%">HS Edit</th>
        <th class="styleTHstd" width="5%">NC User</th>
        <th class="styleTHstd" width="5%">NC Edit</th>
        <th class="styleTHstd" width="5%">MF User</th>
        <th class="styleTHstd" width="5%">MF Edit</th>
        
        <th class="styleTHstd" width="5%">Redo User</th>
        <th class="styleTHstd" width="5%">Redo View</th>
        
        <th class="styleTHstd" width="5%">Admin</th>
        <!--<th class="styleTHstd" width="5%">Dormant</th>-->
        <th class="styleTHstd" width="5%">HS Logs</th>
        <th class="styleTHstd" width="5%">NC logs</th>
        <th class="styleTHstd" width="5%">MF logs</th>
        <th class="styleTHstd" width="5%">Redo logs</th>
    </tr>

    <% 
        While Not UserRs.Eof
        
        UserCount(0) = UserCount(0) +1
        If UserRs("Dormant") = Cbool(True) Then UserCount(1) = UserCount(1) +1
        If UserRs("Password") = "" Then UserCount(2) = UserCount(2) +1
        
        If RowCount = 30 Then
            Response.write "<tr>" & VbCrLf
                Response.write "<td colspan='18' style='height: 50px'>&nbsp;</td>" & VbCrLf
            Response.write "</tr>" & VbCrLf
            Response.write "<tr>" & VbCrLf
            Response.write "<th class='styleTHleft' width='11%'>User</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='10%'>Password</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>Id</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>Status</th>" & VbCrLf        
            Response.write "<th class='styleTHstd' width='5%'>HS User</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>HS View</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>HS Edit</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>NC User</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>NC Edit</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>MF User</th>" & VbCrLf          
            Response.write "<th class='styleTHstd' width='5%'>MF Edit</th>" & VbCrLf
            
            Response.write "<th class='styleTHstd' width='5%'>Redo User</th>" & VbCrLf          
            Response.write "<th class='styleTHstd' width='5%'>Red View</th>" & VbCrLf
            
            
            Response.write "<th class='styleTHstd' width='5%'>Admin</th>" & VbCrLf
            '## Response.write "<th class='styleTHstd' width='5%'>Dormant</th>" & VbCrLf
            
            Response.write "<th class='styleTHstd' width='5%'>HS Logs</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>NC logs</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>MF logs</th>" & VbCrLf
            Response.write "<th class='styleTHstd' width='5%'>Redo logs</th>" & VbCrLf
            
            Response.write "</tr>" & VbCrLf
            RowCount = 0
        End If
   
    %>
    <tr
    <%          
    'If UserRs("Dormant") = Cbool(True)  Then         
    '    Response.write "style='background-color: #FFFF99' title='Dormant User'"             
    'Else
    '    Response.write "style='background-color: #FFFFFF' title='' "
    'End If         
    %> >    
        
        <td class="styleTDBWleft" >
        <%
        If Application("UsersOnLineName" & Cstr(UserRs("Id"))) = "" Then            
        %>            
           <a onmouseover="this.style.color='<%=NewCyan%>'" onmouseout="this.style.color='<%=NewBlue%>'"
           href ="javascript:EditUser('<%Response.Write UserRs("Id")%>');" title="Click to edit" 
           style="color:<%=NewBlue%>; position: relative; top: 1px; "><%Response.Write UserRs("UserName")%></a>
        <%
        Else
            Response.Write "<span title='User online, edit disabled'>" & UserRs("UserName") & "</span>" 
        End If
        %>
        </td>
        
        <td class="styleTDstd"><%=UserRs("Password")%></td>
        <td class="styleTDstd"><%=UserRs("Id")%></td>
        
        <td class="styleTDstd">
        <%
        If Application("UsersOnLineName" & Cstr(UserRs("Id"))) <> "" Then        
            Response.Write "<img align='middle' alt='' title='On line' src='Images/checkmark.ico' width='15' height='15'  />" 
        Else
            If UserRs("Active") = Cbool(False) Then
                Response.Write "&nbsp;"
            Else
                Response.Write "<img align='middle' alt='' title='Off line' src='Images/X2.ico' width='15' height='15'  />"
            End If 
        End If    
        %>
        </td>
         
        <td class="styleTDstd">
        <%        
        If UserRs("ShowHS") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='HS User' src='Images/checkmark.ico' width='15' height='15'  />"
        End If                   
        %>
        </td>
        
        <td class="styleTDstd">
        <%        
        If UserRs("HsView") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='User can view all logs' src='Images/checkmark.ico' width='15' height='15'  />"
        End If                   
        %>
        </td>
        
        <td class="styleTDstd">
        <%
        If UserRs("HsEdit") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='HS Edit' src='Images/checkmark.ico' width='15' height='15'  />"
        End If    
        %>
        </td>
        
        <td class="styleTDstd">
        <%
        If UserRs("ShowNC") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='NC User' src='Images/checkmark.ico' width='15' height='15'  />"
        End If   
        %>
        </td>
        
        <td class="styleTDstd">
        <%
        If UserRs("NcEdit") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='NC Edit' src='Images/checkmark.ico' width='15' height='15'  />"
        End If    
        %>
        </td>
        
        <td class="styleTDstd">
        <%
        If UserRs("ShowMF") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='MF User' src='Images/checkmark.ico' width='15' height='15'  />"
        End If    
        %>
        </td>   
        
        
        <td class="styleTDstd">
        <%
        If UserRs("MfEdit") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='MF Edit' src='Images/checkmark.ico' width='15' height='15'  />"
        End If    
        %>
        </td>
        
        
        <td class="styleTDstd">
        <%        
        If UserRs("RedoAdd") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='Redo User' src='Images/checkmark.ico' width='15' height='15'  />"
        End If                   
        %>
        </td>
        
        <td class="styleTDstd">
        <%        
        If UserRs("RedoView") = Cbool(False) Then
            Response.Write "&nbsp;"
        Else
            Response.Write "<img align='middle' alt='' title='View Logs' src='Images/checkmark.ico' width='15' height='15'  />"
        End If                   
        %>
        </td>
        
        
        
         
        <td class="styleTDstd">
        <%        
            If UserRs("AdminUser") = Cbool(True) Then
                Response.Write "<img align='middle' alt='' title='Admin User' src='Images/checkmark.ico' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td>
        
        <!--<td class="styleTDstd">
        <%        
           ' If UserRs("Dormant") = Cbool(True) Then
           '     Response.Write "<img align='middle' alt='' title='Dormant User' src='Images/Ban.ico' width='15' height='15'  />"
           ' Else
           '     Response.Write "&nbsp;"
           ' End If            
        %>
        </td>-->
        
        <td class="styleTDstd">
            <%= GetLogCount (UserRs("Id"),"HS")%>
        </td>
        
        <td class="styleTDstd">
            <%= GetLogCount (UserRs("Id"),"NC")%>     
        </td>
      
        <td class="styleTDstd">
            <%= GetLogCount (UserRs("Id"),"MF")%>     
        </td>
        
        <td class="styleTDstd">
            <%= GetLogCount (UserRs("Id"),"IHR")%>     
        </td>
        
    </tr>
    <%
        UserRs.MoveNext
        RowCount = RowCount + 1        
    Wend
    UserRs.Close
    Set UserRs = Nothing 
    %>   
    
    <tr>        
        <td  colspan="18" >
            <font size="2" style="color: #0069AA">
                <br />Total Users <%=UserCount(0)%>. <!--Unregistered Users <%'=UserCount(2)%>.--> <!--Dormat Users <%'=UserCount(1)%>.--> 
                HS Logs = <%=LogCount(0)%>. NC Logs = <%=LogCount(1)%>. MF Logs = <%=LogCount(2)%>. Redo Logs = <%=LogCount(3)%> 
            </font>
        </td>    
    </tr>
    <tr>
        <td  colspan="18" style="height: 20px">&nbsp;</td>    
    </tr>

</table>
</body>  
</html>

<%
Function GetLogCount (Uid,LogName)


    Dim LogCountRs
    Dim LogCountSql
    Dim LocalSource
    Dim LocalCount
    LocalCount = 0
    LocalSource = ""
    
    If LogName = "HS" Then
        LocalSource = Session("ConnHSReports")
        LogCountSql = "Select UserId From Logs Where(UserId = " & Uid & ")"
    ElseIf LogName = "NC" Then
        LocalSource = Session("ConnNCReports")
        LogCountSql = "Select UserId From Logs Where(UserId = " & Uid & ")"
    ElseIf LogName = "IHR" Then
        LocalSource = Session("ConnQC")
        LogCountSql = "Select CreatedById From OhQuoteData Where(CreatedById = " & Uid & ")"
    Else
        LocalSource = Session("ConnMachinefaults")
        LogCountSql = "Select UserId From Logs Where(UserId = " & Uid & ")"
    End If
    
    Set LogCountRs = Server.CreateObject("ADODB.Recordset")
    LogCountRs.ActiveConnection = LocalSource  
    LogCountRs.Source = LogCountSql
    LogCountRs.CursorType = Application("adOpenForwardOnly")
    LogCountRs.CursorLocation = Application("adUseClient")
    LogCountRs.LockType = Application("adLockReadOnly")
    LogCountRs.Open
    
    If LogCountRs.BOF = True Or LogCountRs.EOF = True Then
        LocalCount = 0
    Else
        LocalCount = LogCountRs.RecordCount
        If LogName = "HS" Then
            LogCount(0) = LogCount(0) + LogCountRs.RecordCount
        ElseIf LogName = "NC" Then
            LogCount(1) = LogCount(1) + LogCountRs.RecordCount
        ElseIf LogName = "IHR" Then
            LogCount(3) = LogCount(3) + LogCountRs.RecordCount
        Else
            LogCount(2) = LogCount(2) + LogCountRs.RecordCount
        End If
    End If
    
    LogCountRs.Close
    Set LogCountRs = Nothing
    
    GetLogCount = LocalCount
   
End Function   

%>