<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache"
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

Dim UserSql
Dim UserToGet
UserToGet = Request.QueryString("Uid")

'## Get User List
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


'## get data for this user
UserSql = "SELECT * FROM Email Where (Id = " & UserToGet & ")" 

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

<form action="UpdateUser.asp" method="post" name="frmEditEmailUser" id="frmEditEmailUser" onsubmit="return ValidateEmailUser('Edit');" >
<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100"> 
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" /> 
        </td>           
    </tr>
    <tr>
	    <td height="8" valign="top" colspan="3">
	        <hr style="border-style: none;  height: 4px; background-color: <%=MplCyan%>; display: block;" />
	    </td>
    </tr>
    <tr>
        <td align="left" valign="top" width="33%">
            &nbsp;<a href ="javascript:window.location.replace('EmailUsers.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;user&nbsp;page</a>
        </td>
        <td align="center" width="34%">
            <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >Edit&nbsp;Email&nbsp;User</font>
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
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">Email&nbsp;Name</font></td>
            <td valign="middle" width="300px">
            <input name="txtUserID" type="text" id="txtUserID" class="inputboxes" 
                    onmouseup='return false;' onkeyup="TabToNext(event,'txtEmail');"
                onkeypress="return DisableEnterKey(event,'NewName');" 
                    value="<%=UserRs("EmailName")%>" onchange="javascript:UserNameChanged();" 
                size="54" maxlength="48" readonly="readonly"
                />       
            </td>
            <td align="left">&nbsp;</td>
                
        </tr>            
                
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>          
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">Email&nbsp;Address</font></td>
            <td valign="middle" width="300px">
            <input name="txtEmail" type="text" id="txtEmail" class="inputboxes" size="54"
              onmouseup='return false;' onkeypress="return DisableEnterKey(event,'');" 
              value="<%=UserRs("EmailAddress")%>" onchange="EnableSubmit();" maxlength="48" 
                    readonly="readonly"/>       
            </td>
            <td align="left">&nbsp;Although not used it is still required</td>                
        </tr>
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">HS Group</font></td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkHS" name="chkHS" 
                <%If UserRs("InHSGroup") = Cbool(True) Then Response.Write ("checked='checked'")%>  />       
            </td>
            <td align="right">&nbsp;</td>                
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">NC Group</font>&nbsp;</td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkNC" name="chkNC" 
                <%If UserRs("InNCGroup") = Cbool(True) Then Response.Write ("checked='checked'")%>  />       
            </td>
            <td align="right">&nbsp;</td>                
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">MF Group</font>&nbsp;</td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkMF" name="chkMF" 
                <%If UserRs("InMFGroup") = Cbool(True) Then Response.Write ("checked='checked'")%>  />
            </td>
             <td align="right">&nbsp;</td>               
        </tr> 
         
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">MF Critical Group</font>&nbsp;</td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkMFC" name="chkMFC" 
                <%If UserRs("InMFCGroup") = Cbool(True) Then Response.Write ("checked='checked'")%>  />
            </td>
             <td align="right">&nbsp;</td>               
        </tr> 
         
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        
        
        <tr >
            <td valign="middle" width="150px" class="inputlabel" colspan="1"><font size="3">Active</font>&nbsp;</td>
            <td valign="middle" width="300px">
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
            
            <%
            Dim UpdateDisabled
            If UserRs("Isgroup") = CBool(True) Then
                UpdateDisabled = " disabled='disabled' "
            Else
                UpdateDisabled = ""
            End If
            
            %>
            
            <td width="10%">
                <input id="btnSubmit" name="btnSubmit" type="submit" value="Update" <%=UpdateDisabled%> />                   
            </td>
            <td width="10%">
                <input type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            </td>
            <td >
                <input type="<%=InputType %>" name="hUid" id="hUid" value="<%=UserToGet%>" />&nbsp;
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmEditEmailUser" />&nbsp;
                <input type="<%=InputType %>" name="hNameChange" id="hNameChange" value="0"/>&nbsp;
                <input type="<%=InputType %>" name="hUserList" id="hUserList" value="<%=UserList%>"/>               
            </td>            
        </tr>          
    </table>    
    
</form>   
</body>  
</html>

<%
UserRs.Close
Set UserRs = Nothing
%>