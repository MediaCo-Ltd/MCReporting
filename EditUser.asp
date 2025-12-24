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


'## get data for this user
UserSql = "SELECT * FROM Users Where (Id = " & UserToGet & ")" 

Set UserRs = Server.CreateObject("ADODB.Recordset")
UserRs.ActiveConnection = Session("ConnMcLogon")
UserRs.Source = UserSql
UserRs.CursorType = Application("adOpenForwardOnly")
UserRs.CursorLocation = Application("adUseClient")
UserRs.LockType = Application("adLockReadOnly")
UserRs.Open
Set UserRs.ActiveConnection = Nothing

Dim DormantStatus
Dim DelBtnText
Dim DormantHeaderText
Dim HeaderExtra

Dim RetrurnUrl

Dim TableTop
TableTop = "16%"

If UserRs("Dormant") = Cbool(True) Then
    RetrurnUrl = "UsersDormant.asp"
    DormantStatus = "1"
    DelBtnText = "Restore"
    TableTop = "25%"
    DormantHeaderText = "<label style='color: #FF0000'>&nbsp;&nbsp;&nbsp;User is dormant, all items are disabled. Restore user 1st, update & then edit user again</label><br />"
Else
    RetrurnUrl = "UsersOnline.asp"
    DormantStatus = "0"
    DelBtnText = "Delete"
    DormantHeaderText = ""
    TableTop = "16%"
End If

If UserRs("Active") = Cbool(True) Then
    HeaderExtra = "<br />&nbsp;&nbsp;&nbsp;Unticking Active will hide there name from the login drop down. Deleting a user even if not active will remove them from all dropdowns"
Else
    HeaderExtra = ""
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
        function EnableEdit(ChkEdit,ChkId) 
        {
            if (document.getElementById(ChkId).checked == false) 
            {
                document.getElementById(ChkEdit).checked = false;
            }            
        }
    </script>
</head>

<body style="padding: 0px; margin: 0px" >           

<form action="UpdateUser.asp" method="post" name="frmEditUser" id="frmEditUser" onsubmit="return ValidateUser('Edit');" >
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
            &nbsp;<a href ="javascript:window.location.replace('<%=RetrurnUrl%>');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;user&nbsp;page</a>
        </td>
        <td align="center" width="34%">
            <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >Edit&nbsp;User</font>
        </td>
        <td align="right" width="33%">&nbsp;</td>
     </tr>	
</table>
                                                        
                                                        
<table class="NonDataTables" style="width: 100%; position: absolute; top: <%=TableTop%>;" >        
        <tr>
            <td colspan="3">
                <font size="3">                
                <%=DormantHeaderText%>
                <%If DormantStatus = "0" Then%>
                <br />&nbsp;&nbsp;&nbsp;User&nbsp;Name&nbsp;can&nbsp;not&nbsp;be&nbsp;changed.
                <br />&nbsp;&nbsp;&nbsp;Admin&nbsp;users&nbsp;&amp;&nbsp;users&nbsp;who&nbsp;can&nbsp;edit&nbsp;must&nbsp;have&nbsp;a&nbsp;password.
                &nbsp;So&nbsp;they&nbsp;can&nbsp;create&nbsp;their&nbsp;own&nbsp;enter&nbsp;0&nbsp;(zero)&nbsp;in&nbsp;the&nbsp;password&nbsp;box
                <br />&nbsp;&nbsp;&nbsp;Std&nbsp;users&nbsp;don't&nbsp;need&nbsp;a&nbsp;password,&nbsp;so&nbsp;enter&nbsp;anything&nbsp;other&nbsp;than&nbsp;0&nbsp;(zero)&nbsp;in&nbsp;the&nbsp;password&nbsp;box
                <%End If%>                
                <%=HeaderExtra%>
                </font>
            </td>        
        </tr>
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">User&nbsp;Name</font></td>
            <td valign="middle" width="300px">
            <input name="txtUserID" type="text" id="txtUserID" class="inputboxes" onmouseup='return false;' 
                onkeypress='return false;' value="<%=UserRs("UserName")%>"  readonly="readonly" size="54" />       
            </td>
            <td align="left">&nbsp;</td>
                
        </tr>            
                
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>          
        <tr >                                                           
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Password</font></td>
            <td valign="middle" colspan="2">
                <input name="txtPassword" type="text" id="txtPassword" value="<%=UserRs("Password")%>" size="54" maxlength="48"
                class="inputboxes" onfocus='this.select();' onmouseup='return false;' onkeypress="return DisableEnterKey(event,'');"
                <% If DormantStatus = "1" Then Response.Write ("readonly='readonly'")%> />
                &nbsp;<label style="color: #FF0000; padding-bottom: 5px; font-size: 15px;">
                &nbsp;Password&nbsp;can&nbsp;only&nbsp;contain&nbsp;alphanumeric&nbsp;characters&nbsp;plus&nbsp;these&nbsp;£$#@*=+-&nbsp;symbols
                </label>
            </td>
                            
        </tr>
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Active</font></td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkActive" name="chkActive" 
                <%If UserRs("Active") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                 />
                 &nbsp;&nbsp;&nbsp;
                 <!--style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;"-->
                 <input type="button" id="btnDel" name="btnDel" value="<%=DelBtnText%>" onclick="javascript:DelUser();"
                 style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 84px;"
                 />
                  <font size="3" style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 90px;">User</font>      
            </td>
            <td align="right">&nbsp;</td>                
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">HS User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkHS" name="chkHS" onclick="javascript:EnableEdit('ChkHsEdit','chkHS');"
                <%If UserRs("ShowHS") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                />
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">HS Edit</font>&nbsp;&nbsp;
                    <input type="checkbox"  id="ChkHsEdit" name="ChkHsEdit" 
                    <%If UserRs("HsEdit") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                    <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                    />
                </label>
                
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 200px;" title="User can view all HS logs">
                    <font size="3">HS View</font>&nbsp;&nbsp;
                    <input type="checkbox" id="ChkHsView" name="ChkHsView" title="User can view all HS logs"
                    <%If UserRs("HsView") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                    <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>                    
                    />
                </label>
            </td>
            <td align="right">&nbsp;</td>            
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">NC User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkNC" name="chkNC" onclick="javascript:EnableEdit('ChkNcEdit','chkNC');" 
                <%If UserRs("ShowNC") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                <%If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%> 
                />
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">NC Edit</font>&nbsp;&nbsp;
                    <input type="checkbox"  id="ChkNcEdit" name="ChkNcEdit" 
                    <%If UserRs("NcEdit") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                    <%If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                    />
                </label>
            </td>
            <td align="right">&nbsp;</td>            
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">MF User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkMF" name="chkMF" onclick="javascript:EnableEdit('ChkMfEdit','chkMF');" 
                <%If UserRs("ShowMF") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                />
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">MF Edit</font>&nbsp;&nbsp;
                    <input type="checkbox"  id="ChkMfEdit" name="ChkMfEdit" 
                    <%If UserRs("MfEdit") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                    <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                    />
                </label>
                
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 200px;" title="User can edit own record">
                    <font size="3">Edit Own</font>&nbsp;&nbsp;
                    <input type="checkbox" id="ChkEditOwnMf" name="ChkEditOwnMf" 
                    <%If UserRs("UserEditOwnMf") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                    <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                    />
                </label>
                
                
            </td>
            <td align="right">&nbsp;</td>            
        </tr>
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Redo User</font></td>
            <td valign="middle" align="left" >
                <input type="checkbox" id="chkRedo" name="chkRedo" 
                <%If UserRs("RedoAdd") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                <% If DormantStatus = "1" Or Session("HideIndoorRedo") = True Then Response.Write ("disabled='disabled'")%>
                />
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 100px;">
                    <font size="3">Redo View</font>&nbsp;&nbsp;
                    <input type="checkbox"  id="ChkRedoView" name="ChkRedoView"
                    <%If UserRs("RedoView") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                    <% If DormantStatus = "1" Or Session("HideIndoorRedo") = True Then Response.Write ("disabled='disabled'")%>               
                    />  
                </label>
                
                <label style="color: #0069AA; vertical-align: text-bottom; position: relative; left: 175px;" title="User can see redo cost">
                    <font size="3">Redo Cost</font>&nbsp;&nbsp;
                    <input type="checkbox" id="ChkRedoCost" name="ChkRedoCost" 
                    <%If UserRs("ShowRedoCost") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                    <% If DormantStatus = "1" Or Session("HideIndoorRedo") = True Then Response.Write ("disabled='disabled'")%>
                    />
                </label>
            </td>
            <td align="right">&nbsp;</td>            
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Admin User</font></td>
            <td valign="middle" colspan="2">
                <input type="checkbox" id="chkAdmin" name="chkAdmin" 
                <%If UserRs("AdminUser") = Cbool(True) Then Response.Write ("checked='checked'")%>  
                <% If DormantStatus = "1" Then Response.Write ("disabled='disabled'")%>
                /> &nbsp;<font size="3">Admin Users must have a password</font>    
            </td>
                           
        </tr> 
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
               
        <tr>
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Email</font></td>
            <td valign="middle"  colspan="2">
            <input type="text" id="txtEmail" name="txtEmail" maxlength="50" size="54" value="<%=UserRs("PwEmail")%>" />
            <font size="3">Only for users who must have a password</font> 
            </td>
        </tr> 
 
</table>
                                                        
<table style="width: 100%; position: absolute; bottom: 11%; padding-right: 20px; padding-left: 20px;">
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
                <input id="btnSubmit" name="btnSubmit" type="submit" value="Update" />                   
            </td>
            <td width="10%">
                <input type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            </td>
            <td >
                <input type="<%=InputType %>" name="hUid" id="hUid" value="<%=UserToGet%>" />&nbsp;
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmEditUser" />&nbsp;
                <input type="<%=InputType %>" name="hNameChange" id="hNameChange" value="0"/>
                <input type="<%=InputType %>" name="hDelete" id="hDelete" value="<%=DormantStatus%>"/>              
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