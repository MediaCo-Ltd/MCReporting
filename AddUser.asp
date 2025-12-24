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
UserRs.Source = "SELECT * FROM Users Order By Id"
UserRs.CursorType = Application("adOpenForwardOnly")
UserRs.CursorLocation = Application("adUseClient")
UserRs.LockType = Application("adLockReadOnly")
UserRs.Open
Set UserRs.ActiveConnection = Nothing

While Not UserRs.EOF
    If UserList = "" Then    
        UserList = Ucase(UserRs("UserName"))
    Else
        UserList = UserList & "#" & Ucase(UserRs("UserName"))
    End If
    UserRs.MoveNext
Wend

UserRs.Close
Set UserRs = Nothing

Dim RedirectUrl
Dim RedirectTxt

If Request.QueryString("rt") = "1" Then
    RedirectUrl = "Admin.asp"
    RedirectTxt = "Return&nbsp;to&nbsp;admin&nbsp;page"
Else
    RedirectUrl = "UsersOnLine.asp"
    RedirectTxt = "Return&nbsp;to&nbsp;user&nbsp;page"
End If



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
    <script type="text/javascript" >
        function EnableEdit(ChkId) 
        {
            if (document.getElementById(ChkId).disabled == true) 
            {
                document.getElementById(ChkId).disabled = false;
            }
            else 
            {
                document.getElementById(ChkId).disabled = true;                
                document.getElementById(ChkId).checked = false;
            }
        }
    </script>
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:UserLoad();" >           
<form action="UpdateUser.asp" method="post" name="frmAddUser" id="frmAddUser" onsubmit="return ValidateUser('Add');" >
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
                &nbsp;<a href ="javascript:window.location.replace('<%=RedirectUrl %>');" style="font-size:12px; color:<%=NewBlue%>"><%=RedirectTxt%></a>
            </td>
            <td align="center" width="34%">
                <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
                <font style="color: <%=NewBlue%>; font-weight: bold; font-size: 16px;" >Add&nbsp;New&nbsp;User</font>
            </td>            
            <td align="right" width="33%">&nbsp;</td>
        </tr>                	        
    </table>
    
    <table class="NonDataTables" style="width: 100%; position: absolute; top: 18%;" >
        
        <tr>
            <td colspan="2">
                <font size="3">
                &nbsp;&nbsp;&nbsp;Admin&nbsp;users&nbsp;&amp;&nbsp;users&nbsp;who&nbsp;can&nbsp;edit&nbsp;must&nbsp;have&nbsp;a&nbsp;password.
                <br />&nbsp;&nbsp;&nbsp;So&nbsp;they&nbsp;can&nbsp;create&nbsp;their&nbsp;own&nbsp;enter&nbsp;0&nbsp;(zero)&nbsp;in&nbsp;the&nbsp;password&nbsp;box
                <br />&nbsp;&nbsp;&nbsp;Std&nbsp;users&nbsp;don't&nbsp;need&nbsp;a&nbsp;password,&nbsp;so&nbsp;enter&nbsp;anything&nbsp;other&nbsp;than&nbsp;0&nbsp;(zero)&nbsp;in&nbsp;the&nbsp;password&nbsp;box
                <br />&nbsp;&nbsp;&nbsp;Selecting&nbsp;HS&nbsp;Edit&nbsp;or&nbsp;HS&nbsp;View&nbsp;will&nbsp;require&nbsp;a&nbsp;password.&nbsp;NC&nbsp;is&nbsp;now&nbsp;disabled.
                </font>            
            </td>        
        </tr>
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">User&nbsp;Name</font></td>
            <td valign="middle" >
            <input name="txtUserID" type="text" id="txtUserID" class="inputboxes" onfocus='this.select();' onmouseup='return false;' 
                onkeypress="return DisableEnterKey(event,'NewName');" onkeyup="TabToNext(event,'txtPassword');" value="" onchange="EnableSubmit();" 
                size="54" maxlength="48"
                />       
            </td>
                 
        </tr>            
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>        
        <tr >                                                           
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Password</font></td>
            <td valign="middle" >
                <input name="txtPassword" type="text" id="txtPassword" value="" class="inputboxes" size="54" maxlength="48"
                onfocus='this.select();' onmouseup='return false;' onkeypress="return DisableEnterKey(event,'');" />
                &nbsp;<label style="color: #FF0000; padding-bottom: 5px; font-size: 15px;">
                &nbsp;Password&nbsp;can&nbsp;only&nbsp;contain&nbsp;alphanumeric&nbsp;characters&nbsp;plus&nbsp;these&nbsp;£$#@*=+-&nbsp;symbols
                </label>
            </td>              
        </tr>
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>        
       
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Active</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkActive" name="chkActive"  checked="checked" />
                &nbsp;<font size="3">By default new users are active</font>         
            </td>            
        </tr> 
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">HS User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkHS" name="chkHS" onclick="javascript:EnableEdit('ChkHsEdit');" />
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">HS Edit</font>&nbsp;&nbsp;
                    <input type="checkbox" id="ChkHsEdit" name="ChkHsEdit" disabled="disabled" />
                </label>
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 200px;" title="User can view all HS logs">
                    <font size="3">HS View</font>&nbsp;&nbsp;
                    <input type="checkbox" id="ChkHsView" name="ChkHsView" title="User can view all HS logs"/>
                </label>
            </td>            
        </tr> 
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">NC User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkNC" name="chkNC" onclick="javascript:EnableEdit('ChkNcEdit');" disabled="disabled" />
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">NC Edit</font>&nbsp;&nbsp;
                    <input type="checkbox"  id="ChkNcEdit" name="ChkNcEdit" disabled="disabled"/>
                </label>
            </td>            
        </tr> 
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">MF User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkMF" name="chkMF" onclick="javascript:EnableEdit('ChkMfEdit');"/>
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">MF Edit</font>&nbsp;&nbsp;
                    <input type="checkbox"  id="ChkMfEdit" name="ChkMfEdit" disabled="disabled"/>
                </label>
            </td>            
        </tr> 
        
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Redo User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkRedo" name="chkRedo" <%If Session("HideIndoorRedo") = True Then Response.Write ("disabled='disabled'") %> />
                  <!--onclick="javascript:EnableEdit('ChkRedoView');"-->
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">Redo View</font>&nbsp;&nbsp;
                    <input type="checkbox"  id="ChkRedoView" name="ChkRedoView" <%If Session("HideIndoorRedo") = True Then Response.Write ("disabled='disabled'") %> />  
                    <!--disabled="disabled"-->
                </label>
                
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 175px;" title="User can see redo cost">
                    <font size="3">Redo Cost</font>&nbsp;&nbsp;
                    <input type="checkbox" id="ChkRedoCost" name="ChkRedoCost" title="User can see redo cost" <%If Session("HideIndoorRedo") = True Then Response.Write ("disabled='disabled'") %> />
                </label>
            </td>            
        </tr>
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Admin User</font></td>
            <td valign="middle" align="left" >
            <input type="checkbox" id="chkAdmin" name="chkAdmin"  />
            &nbsp;<font size="3">Admin Users must have a password</font>  
            </td>            
        </tr> 
        
        
        <tr>
            <td height="20px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Email</font></td>
            <td valign="middle" align="left" >
            <input type="text" id="txtEmail" name="txtEmail" maxlength="50" size="54" value="" />
            &nbsp;<font size="3">Only for users who must have a password</font>  
            </td>            
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
                
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmAddUser" />&nbsp;
                <input type="<%=InputType %>" name="hUserList" id="hUserList" value="<%=UserList%>"/>&nbsp;
                <input type="<%=InputType %>" name="hNameChange" id="hNameChange" value="1"/>&nbsp;
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