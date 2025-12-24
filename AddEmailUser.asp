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

If Session("ConnMcLogon") = "" Then Response.Redirect("Admin.asp")

Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"
End If

Dim UserRs
Dim UserList

Set UserRs = Server.CreateObject("ADODB.Recordset")
UserRs.ActiveConnection = Session("ConnMcLogon")
UserRs.Source = "SELECT Id, EmailName FROM Email Order By Id"
UserRs.CursorType = Application("adOpenForwardOnly")
UserRs.CursorLocation = Application("adUseClient")
UserRs.LockType = Application("adLockReadOnly")
UserRs.Open
Set UserRs.ActiveConnection = Nothing

While Not UserRs.EOF
    If UserList = "" Then    
        UserList = Ucase(UserRs("EmailName"))
    Else
        UserList = UserList & "#" & Ucase(UserRs("EmailName"))
    End If
    UserRs.MoveNext
Wend

UserRs.Close
Set UserRs = Nothing

Dim NewUserRs

Set NewUserRs = Server.CreateObject("ADODB.Recordset")
NewUserRs.ActiveConnection = Session("ConnMcLogon")
NewUserRs.Source = "SELECT Id, UserName FROM Users Where (ActiveEmail = 0) And (Active = 1) Order By UserName"
NewUserRs.CursorType = Application("adOpenForwardOnly")
NewUserRs.CursorLocation = Application("adUseClient")
NewUserRs.LockType = Application("adLockReadOnly")
NewUserRs.Open
Set NewUserRs.ActiveConnection = Nothing


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/PenPad64.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/MCReportsJSFunc.js"></script>   
</head>

<body style="padding: 0px; margin: 0px" >           
<form action="UpdateUser.asp" method="post" name="frmAddEmailUser" id="frmAddEmailUser" onsubmit="return ValidateEmailUser('Add');" >
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
            <td align="left" valign="top" width="33%" >
                &nbsp;<a href ="javascript:window.location.replace('EmailUsers.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;user&nbsp;page</a>
            </td>
            <td align="center" width="34%">
                <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
                <font style="color: <%=NewBlue%>; font-weight: bold; font-size: 16px;" >Add&nbsp;User&nbsp;to&nbsp;Email&nbsp;Group</font>
            </td>            
            <td align="right" width="33%">&nbsp;</td>
        </tr>                	        
    </table>
    
    <table class="NonDataTables" style="width: 100%; position: absolute; top: 28%;" >
        
        <tr>
            <td colspan="3">
                <font size="3">
                &nbsp;&nbsp;All&nbsp;emails&nbsp;are&nbsp;now&nbsp;sent&nbsp;via&nbsp;groups.<br />
                &nbsp;&nbsp;This&nbsp;is&nbsp;just&nbsp;to&nbsp;show&nbsp;who&nbsp;is&nbsp;in&nbsp;which&nbsp;group.
                Users&nbsp;need&nbsp;to&nbsp;be&nbsp;added&nbsp;to&nbsp;groups&nbsp;in&nbsp;Exchange&nbsp;on&nbsp;MC-Exch
                </font>            
            </td>        
        </tr>
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">Email&nbsp;Name</font></td>
            <td valign="middle" >
            <select id="cboEmailUser" onchange="javascript:PickEmailUser()">
                    <option value="" >Select User</option>
                    <%
                    While Not NewUserRs.EOF
                        Response.Write "<option value=" & NewUserRs("Id") & ">" & NewUserRs("UserName") & "</option>"
                        NewUserRs.MoveNext
                    Wend
                    
                    NewUserRs.Close
                    Set NewUserRs = Nothing
                    %>
                
                </select>
            
                  
            </td>
            <td align="left">&nbsp;</td>    
        </tr>            
       
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr> 
               
        <tr >                                                           
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">Email&nbsp;Address</font></td>
            <td valign="middle" >
                <input name="txtEmail" type="text" id="txtEmail" value="" class="inputboxes" onkeyup="javascript:EnableSubmit('');"
                onfocus='this.select();' onmouseup='return false;' onkeypress="return DisableEnterKey(event,'');" size="54" maxlength="48"  />
            </td>
            <td align="left">&nbsp;Although not used it is still required to enable update</td>  
        </tr>
        
        
         <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">HS Group</font></td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkHS" name="chkHS" />       
            </td>
            <td align="right">&nbsp;</td>               
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">NC Group</font></td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkNC" name="chkNC" />       
            </td>
            <td align="right">&nbsp;</td>               
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">MF Group</font></td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkMF" name="chkMF"  />       
            </td>
            <td align="right">&nbsp;</td>                
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">MF Critical Group</font>&nbsp;</td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkMFC" name="chkMFC" />
            </td>
             <td align="right">&nbsp;</td>               
        </tr> 
         
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        
        
        
        <tr>
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">Active</font></td>
            <td valign="middle" align="left" >
            <input type="checkbox" id="chkActive" name="chkActive" disabled="disabled" />
            </td> 
            <td align="right">&nbsp;</td>           
        </tr> 
                        
    </table>
    
    <table style="width: 100%; position: absolute; bottom: 12%; padding-right: 20px; padding-left: 20px;">
        <tr>
            <td width="10%">&nbsp;
                <%
                    If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                        Response.Write ("Don't Save Data&nbsp;&nbsp;")
                        Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                    End If          
                 %>            
            </td>
            <td width="10%">
                <input id="btnSubmit" name="btnSubmit" type="submit" value="Update" disabled="disabled" />                   
            </td>
            <td width="10%">
                <input type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            </td>
            <td >
                
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmAddEmailUser" />
                <input type="<%=InputType %>" name="hUserList" id="hUserList" value="<%=UserList%>"/>
                <input type="<%=InputType %>" name="hNameChange" id="hNameChange" value="1"/> 
                <input type="<%=InputType %>" name="hEmailUserId" id="hEmailUserId" value=""/>
                <input type="<%=InputType%>" name="txtUserID" id="txtUserID" value=""  /> 
             
                
                
            </td>            
        </tr>          
    </table>
    <br />
    <br />            
    <table style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
        <tr>  
            <td height="50" >
                <hr style="width: 98%; border-style: none;  height: 1px; background-color: <%=NewCyan%>; display: block;" />
                <p align="center"><font size="2"> MediaCo Ltd. Churchill Point, Churchill Way, 
                    Trafford Park, Manchester M17 1BS</font>
                </p>
            </td>
        </tr>
    </table>
</form>
</body>  
</html>