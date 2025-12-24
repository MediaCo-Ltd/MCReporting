<%@language="vbscript" codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\PkId.asp" -->

<%
If Session("UserName") = "" Then Response.Redirect ("Login.asp")

Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

Response.expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"


Dim ErrMsg
Dim LocalJobType
Dim ReturnAddress

ErrMsg = Session("JobNoError")

Session("JobNo") = ""
Session("JobNoError") = ""

Dim TheDay
TheDay = WeekdayName(Weekday(Date,1) ,False,1)

If Session("ShowText") = CBool(True) Then
    Session("ShowHidden") = "text"
    Session("ShowHiddenSize") = " size= '3' "
Else
    Session("ShowHidden") = "hidden"
    Session("ShowHiddenSize") = ""
End If

Dim Disabled
If Session("AddRedo") = CBool(False) Then 
    Disabled = " disabled='disabled' "
Else
    Disabled = ""
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
    LockRs.Source = LockRs.Source & " And (SystemName = 'RD')"
    LockRs.CursorType = Application("adOpenStatic")
    LockRs.CursorLocation = Application("adUseClient")
    LockRs.LockType = Application("adLockOptimistic")
    LockRs.Open

    If LockRs.BOF = True Or LockRs.EOF = True Then
        '## No match so do nothing
    Else
        If LockRs("LockedbyId") = Session("UserId") Then  'And LockRs("SystemName") = "HS"
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

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>In House Redo</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/IhRedoJSFunc.js"></script>
    <link rel="shortcut icon" href="Images/X1.ico" type="image/x-icon" />    
</head>

<body style="padding: 0px; margin: 0px" onload="PageLoadChk();">

<form action="IhrCheckJobNo.asp" method="post" name="frmJobNo" id="frmJobNo" onsubmit="return ValidateRedoJobNo();">

<table  style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
		<td align="left" valign="bottom" height="100" colspan="3">
            <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
        </td>        
	</tr>
	<tr>
		<td height="8" valign="top" colspan="3">
		    <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		</td>
	</tr>
	
	<tr>
        <td width="33%" style="height: 20px">&nbsp;&nbsp;
        <a id="logoffL" href ="javascript:window.location.replace('SelectSite.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return&nbsp;to&nbsp;select&nbsp;system</a></td>
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/X1.ico" style="width: 20px; height: 20px;" />
            <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">In House Redo&nbsp;(<%=Session("UserName")%>)</font>
        </td>
        <td align="right" width="33%"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
    </tr>

</table> 


<table id="subsection1" style="width: 100%; position: absolute; top: 41%; padding-right: 10px; padding-left: 10px;" >
    <tr>   
        <td >
            <p >  <!--align="center"-->
            &nbsp;&nbsp;&nbsp;&nbsp;<font size="5" color="red">
            Make&nbsp;sure&nbsp;you&nbsp;have&nbsp;put&nbsp;your&nbsp;operation&nbsp;on&nbsp;hold&nbsp;in&nbsp;Clarity&nbsp;before&nbsp;you&nbsp;create&nbsp;the&nbsp;record.			
            </font></p>
            <br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="3" color="<%=NewBlue%>">To create new record. Enter job No (Just number, no REF) then press enter or click Submit</font></p>
            <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="txtJobNo" id="txtJobNo" style="color: black;" <%=Disabled%> />
            &nbsp;<input name="ErrBox" id="ErrBox" type="hidden" value="<%=ErrMsg%>" />
            </p>			
        </td>
    </tr>    
</table>

<table id="subsection2" style="width: 100%; position: absolute; top: 72%; padding-right: 10px; padding-left: 10px;" >
    <tr>   
        <td >
            <input style="position: relative; left: 30px;" type="submit" value="Submit" <%=Disabled%>/>
            &nbsp;&nbsp;&nbsp;
            <%If Session("ViewRedo") = Cbool(True) Then%>
            <input style="position: relative; left: 90px;" type="button" value="View Records"  onclick="javascript:ViewData();"/>
            <%End If%>
        </td>
    </tr>
</table>

<table  style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
    <tr>  
        <td height="50" >
            <hr style="border-style: none; height: 1px; background-color: <%=NewCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
            Trafford Park, Manchester M17 1BS <br />Tel:(+44)161 875 2020 Fax:(+44)161 873 7740</font></p>
        </td>
    </tr>
</table>
</form>
<span id="loading" style="font-size: medium; color:<%=NewBlue%>; font-weight: bold; position: absolute; top: 45%; left: 43%;"></span>
</body> 
</html>
 