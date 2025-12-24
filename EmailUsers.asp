<%@language="vbscript" codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<% 
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"
If Session("ConnMcLogon") = "" Then Response.Redirect("Admin.asp")

Dim EditImg
Dim UserRs
Dim UserSql
Dim ShowhAction

UserSql = "Select Id, EmailName, EmailAddress, EmailActive, EmailHS, EmailMF, EmailNC," 
UserSql = UserSql & " IsGroup, InHSGroup, InMFGroup ,InNCGroup, InMFCGroup from Email Where (EmailAddress <> '') Order by IsGroup DESC, EmailName"

    
Set UserRs = Server.CreateObject("ADODB.Recordset")
UserRs.ActiveConnection = Session("ConnMcLogon")
UserRs.Source = UserSql
UserRs.CursorType = Application("adOpenForwardOnly")        
UserRs.CursorLocation = Application("adUseClient")   
UserRs.LockType = Application("adLockReadOnly")         
UserRs.Open 
Set UserRs.ActiveConnection = Nothing

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
            &nbsp;<a href ="javascript:window.location.replace('Admin.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;admin&nbsp;page</a>
        </td>
	    <td  align="center" width="34%">
	        <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >MediaCo&nbsp;Reporting&nbsp;Email&nbsp;Recipient</font>
	    </td>
	    <td align="right" width="33%">&nbsp;</td>
	</tr>
	
	
	<tr>
	    <td  align="center" valign="top" colspan="3">
	    <br />
	        <font style="font-weight: bold; color:black; font-size: 12px;" >
	        All emails are sent via Groups, this page is only to show who is in which. Users need to be added to group in Exchange on MC-Exch
	        </font>    
	    </td>
    </tr>   
</table>

<table style="width: 80%; position: absolute; top: 30%; left: 10%;" cellpadding="0" cellspacing="0" >
    <tr>
        <th class="styleTHleft" width="15%">Email Name</th>
        <th class="styleTHstd" width="15%">Email Address</th>
        <th class="styleTHstd" width="8%">Email HS</th>
        <th class="styleTHstd" width="8%">Email NC</th>
        <th class="styleTHstd" width="8%">Email MF</th> 
        <th class="styleTHstd" width="8%">In HS Group</th>
        <th class="styleTHstd" width="8%">In MF Group</th>
        <th class="styleTHstd" width="8%">In MF Critical Group</th>
        <th class="styleTHstd" width="8%">In NC Group</th>       
    </tr>

    <% While Not UserRs.Eof%>
    <tr>
        <td class="styleTDBWleft" >
            <%
                If UserRs("Isgroup") = CBool(True) Then
            %>
                <label style="color:black; position: relative; top: 1px; font-weight: bold;">
                    <%Response.Write UserRs("EmailName")%>
                </label>
            
            <%
                Else
            %>
               <a onmouseover="this.style.color='<%=NewCyan%>'" onmouseout="this.style.color='<%=NewBlue%>'"
               href ="javascript:EditEmailUser('<%Response.Write UserRs("Id")%>');" title="Click to edit" 
               style="color:<%=NewBlue%>; position: relative; top: 1px; "><%Response.Write UserRs("EmailName")%></a>
            <%
                End If
            %>
        </td>
        
        <td class="styleTDstd"><%=UserRs("EmailAddress")%></td>
        
        <td class="styleTDstd">
        <%        
            If UserRs("EmailHS") = Cbool(True) Then
                Response.Write "<img align='middle' alt=''  src='Images/checkmark.ico' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td>
        
        <td class="styleTDstd">
        <%        
            If UserRs("EmailNC") = Cbool(True) Then
                Response.Write "<img align='middle' alt=''  src='Images/checkmark.ico' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td>
        
        <td class="styleTDstd">
        <%        
            If UserRs("EmailMF") = Cbool(True) Then
                Response.Write "<img align='middle' alt=''  src='Images/checkmark.ico' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td>
                      
        <td class="styleTDstd">
        <%        
            If UserRs("InHSGroup") = Cbool(True) Then
                Response.Write "<img align='middle' alt='' title='User In Group' src='Images/Plus-icon.png' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td> 
        
        <td class="styleTDstd">
        <%        
            If UserRs("InMFGroup") = Cbool(True) Then
                Response.Write "<img align='middle' alt='' title='User In Group' src='Images/W-S48.png' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td>
        
         <td class="styleTDstd">
        <%        
            If UserRs("InMFCGroup") = Cbool(True) Then
                Response.Write "<img align='middle' alt='' title='User In Group' src='Images/W-S48.png' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td>  
        
        <td class="styleTDstd">
        <%        
            If UserRs("InNCGroup") = Cbool(True) Then
                Response.Write "<img align='middle' alt='' title='User In Group' src='Images/warning.png' width='15' height='15'  />"
            Else
                Response.Write "&nbsp;"
            End If            
        %>
        </td> 
        
    </tr>
    <%
        UserRs.MoveNext        
    Wend
    UserRs.Close
    Set UserRs = Nothing 
    %>
    <tr>
        <td height="10px" colspan="3">&nbsp;</td>    
    </tr>
    <tr>
        <td height="20px" colspan="3" align="left">&nbsp;
            <font size="2" style="color: #0069AA">Click&nbsp;Name&nbsp;to&nbsp;edit&nbsp;user.</font>
        </td>
           
    </tr>
</table>

<table style="width: 100%; position: absolute; bottom: 15%; padding-right: 20px; padding-left: 20px;">
    <tr>
        <td width="10%">&nbsp;</td>
        <td width="15%">
            <input id="btnAdd" type="button" value="Add User" onclick="javascript:window.location.replace('AddEmailUser.asp');"/>
        </td>
        <td width="60%">&nbsp;</td>                 
    </tr>
</table>
              
<table  style="width: 100%; position: absolute; bottom: 5px;  padding-right: 10px; padding-left: 10px;">
    <tr>  
        <td height="50" >
            <hr style="width: 98%; border-style: none;  height: 1px; background-color: <%=NewCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
                Trafford Park, Manchester M17 1BS</font></p>
        </td>
    </tr>
</table>

</body>  
</html>