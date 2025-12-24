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

If Session("ConnMachinefaults") = "" Then Response.Redirect("Admin.asp")

Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"
End If

Dim Jump
Dim Count
Dim SubCount
Dim RowCount
Dim ColCount
Dim Remainder
Dim MachineData()
Dim MachineRs
Dim MachineList
Dim MachineCount


'## Get Machine Names
Set MachineRs = Server.CreateObject("ADODB.Recordset")
MachineRs.ActiveConnection = Session("ConnMachinefaults")
MachineRs.Source = "SELECT * From Machine Order By Id"
MachineRs.CursorType = Application("adOpenForwardOnly")
MachineRs.CursorLocation = Application("adUseClient")
MachineRs.LockType = Application("adLockReadOnly")
MachineRs.Open
Set MachineRs.ActiveConnection = Nothing

MachineCount = MachineRs.RecordCount 
Redim MachineData(MachineCount,3)

'## Poulate Name in Element 0
While Not MachineRs.EOF
    MachineData(MachineRs.AbsolutePosition,0) = Cstr(MachineRs("MachineName"))
       
    If MachineRs("Active") = Cbool(True) Then
        MachineData(MachineRs.AbsolutePosition,1) = "checked='checked'"
    Else
        MachineData(MachineRs.AbsolutePosition,1) = ""
    End If
    
    MachineData(MachineRs.AbsolutePosition,2) = Cstr(MachineRs("Id"))
    MachineData(MachineRs.AbsolutePosition,3) = Cstr(MachineRs("Id"))
    

    MachineRs.MoveNext
Wend

MachineRs.Close
Set MachineRs = Nothing

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>Machine Faults Admin</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/MachineFaultCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/MachineFaultExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/W-S48.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/MachineFaultJSFunc.js"></script>
    <script type="text/javascript" src="JsFiles/MachineFaultAjaxFunc.js"></script>    
</head>

<body style="padding: 0px; margin: 0px">
           
<form action="MfUpdateMachine.asp" method="post" name="frmMachineStatus" id="frmMachineStatus"  >
<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100"> 
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
            &nbsp;<a href ="javascript:window.location.replace('MfAdmin.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;MF&nbsp;admin&nbsp;page</a>
        </td>
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" />
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >Edit&nbsp;Machine</font>
        </td>
        <td align="right" width="33%">&nbsp;</td>
     </tr>	
</table>

<table style="width: 100%; position: absolute; top: 20%; padding-right: 10px; padding-left: 20px;">
    <tr>
        <td height="25px" valign="middle" colspan="6" >
            <p><font size="3" style="color: #0069AA">
                Tick&nbsp;to&nbsp;enable,&nbsp;un-tick&nbsp;to&nbsp;disable&nbsp;machine.
            </font></p>
        </td>
    </tr>
    <tr>
	    <td colspan="6">&nbsp;</td>
    </tr> 
    <%
        '## Does products in blocks of 6, so if 26 products does 4 rows 6 coloumns, remainder of 2 is done in next section
        ColCount = 1        
        Remainder = MachineCount Mod 6
        RowCount = MachineCount\6  
        Jump = RowCount * 6
        For Count = 1 To RowCount     
    %>
    
    <tr>
        <td class="supplierlabel" width="16%"><font size="2"><%=MachineData(ColCount,0)%></font></td>
        <td class="supplierlabel" width="16%"><font size="2"><%=MachineData(ColCount+1,0)%></font></td>
        <td class="supplierlabel" width="16%"><font size="2"><%=MachineData(ColCount+2,0)%></font></td>
        <td class="supplierlabel" width="16%"><font size="2"><%=MachineData(ColCount+3,0)%></font></td>
        <td class="supplierlabel" width="16%"><font size="2"><%=MachineData(ColCount+4,0)%></font></td>
        <td class="supplierlabel" width="16%"><font size="2"><%=MachineData(ColCount+5,0)%></font></td>
    </tr>
    <tr>
        <td class="supplierinput" nowrap="nowrap">Active
            <input type="checkbox" id="Active<%=Cstr(ColCount)%>" name="Active<%=Cstr(ColCount)%>" style="width: 20px" 
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' <%=MachineData(ColCount,1) %> />&nbsp;Id&nbsp;       
        
            <input type="text" id="Sort<%=Cstr(ColCount)%>" name="Sort<%=Cstr(ColCount)%>" style="width: 20px" onfocus="this.select();"
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' readonly="readonly" value="<%=MachineData(ColCount,3) %>" />
       </td>        
        
        <td class="supplierinput" nowrap="nowrap">Active
            <input type="checkbox" id="Active<%=Cstr(ColCount+1)%>" name="Active<%=Cstr(ColCount+1)%>"  style="width: 20px"  
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' <%=MachineData(ColCount+1,1)%> />&nbsp;Id&nbsp;
            
            <input type="text" id="Sort<%=Cstr(ColCount+1)%>" name="Sort<%=Cstr(ColCount+1)%>"  style="width: 20px" onfocus="this.select();" 
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' readonly="readonly" value="<%=MachineData(ColCount+1,3)%>" />
        </td>
        
        <td class="supplierinput" nowrap="nowrap">Active
            <input type="checkbox" id="Active<%=Cstr(ColCount+2)%>" name="Active<%=Cstr(ColCount+2)%>" style="width: 20px" 
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' <%=MachineData(ColCount+2,1)%>/>&nbsp;Id&nbsp;
            
            <input type="text" id="Sort<%=Cstr(ColCount+2)%>" name="Sort<%=Cstr(ColCount+2)%>" style="width: 20px" onfocus="this.select();"
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' readonly="readonly" value="<%=MachineData(ColCount+2,3)%>" /> 
        </td>
        
        <td class="supplierinput" nowrap="nowrap">Active
            <input type="checkbox" id="Active<%=Cstr(ColCount+3)%>" name="Active<%=Cstr(ColCount+3)%>" style="width: 20px" 
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' <%=MachineData(ColCount+3,1)%> />&nbsp;Id&nbsp;
            
            <input type="text" id="Sort<%=Cstr(ColCount+3)%>" name="Sort<%=Cstr(ColCount+3)%>" style="width: 20px" onfocus="this.select();"
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' readonly="readonly" value="<%=MachineData(ColCount+3,3)%>" />
        </td>
        
        <td class="supplierinput" nowrap="nowrap">Active
            <input type="checkbox" id="Active<%=Cstr(ColCount+4)%>" name="Active<%=Cstr(ColCount+4)%>" style="width: 20px" 
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' <%=MachineData(ColCount+4,1)%> />&nbsp;Id&nbsp;
            
            <input type="text" id="Sort<%=Cstr(ColCount+4)%>" name="Sort<%=Cstr(ColCount+4)%>" style="width: 20px" onfocus="this.select();"
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' readonly="readonly" value="<%=MachineData(ColCount+4,3)%>" />
        </td>
        
        <td class="supplierinput" nowrap="nowrap">Active
            <input type="checkbox" id="Active<%=Cstr(ColCount+5)%>" name="Active<%=Cstr(ColCount+5)%>" style="width: 20px" 
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' <%=MachineData(ColCount+5,1)%> />&nbsp;Id&nbsp;
                        
            <input type="text" id="Sort<%=Cstr(ColCount+5)%>" name="Sort<%=Cstr(ColCount+5)%>" style="width: 20px" onfocus="this.select();"
            onmouseup='return false;' onkeypress='return DisableEnterKey(event);' readonly="readonly" value="<%=MachineData(ColCount+5,3)%>" />            
        </td>       
        
    </tr>
    <tr>
        <td  colspan="6" height="5px"></td>      
    </tr>
    
     <%  
            ColCount = ColCount + 6
            If ColCount > Jump Then Exit For
        Next
        
        '## Do any extra products if any to do
        '## If product count \6 leaves no remainder next section will not be created
        If Remainder > 0 Then
            ColCount = Jump +1
            Response.Write "<tr>" & VbCrLf
            For SubCount = Jump +1 To MachineCount     
                Response.Write "<td class='supplierlabel' ><font size='2'>" & MachineData(SubCount,0) & "</font></td>" & VbCrLf
            Next
            Response.Write "</tr>" & VbCrLf
                 
            Response.Write "<tr>" & VbCrLf            
            For SubCount = Jump +1 To MachineCount
                Response.Write "    <td class='supplierinput' nowrap='nowrap'>Active" & VbCrLf        
                Response.Write "        <input type='checkbox' id='Active" & Cstr(SubCount) & "' name='Active" & Cstr(SubCount) & "' style='width: 20px;' " & VbCrLf
                Response.Write "        onmouseup='return false;' onkeypress='return DisableEnterKey(event);' " & MachineData(SubCount,1) & " />&nbsp;Id&nbsp;" & VbCrLf
                Response.Write "" & VbCrLf
                Response.Write "        <input type='text' id='Sort" & Cstr(SubCount) & "' name='Sort" & Cstr(SubCount) & "' style='width: 20px;' onfocus='this.select();'" & VbCrLf
                Response.Write "        onmouseup='return false;'  onkeypress='return DisableEnterKey(event);'  readonly='readonly' value='" & MachineData(SubCount,3) & "'/>" & VbCrLf
                Response.Write "    </td>" & VbCrLf
            Next 
            
            Response.Write "</tr>" & VbCrLf   
        End If
    %>
    
    <tr>
    <td colspan="6" height="3px"><br /></td>
 </tr> 

<tr>
    <td >&nbsp;
        <%
            If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                Response.Write ("Don't Save Data&nbsp;&nbsp;")
                Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
            End If          
         %>            
    </td>
    <td >
        <input id="btnSubmit" name="btnSubmit" type="submit" value="Update" />                   
    </td>
    <td >
        <input type="reset" value="Reset" onclick="javascript:ResetPage();"/>
    </td>        
    <td colspan="3" >                
        <input type="<%=InputType %>" name="frmName" id="frmName" value="frmMachineStatus" />&nbsp;&nbsp;            
        <input type="<%=InputType %>" name="hPrCount" id="hPrCount" value="<%=MachineCount%>"/>
        <input type="<%=InputType %>" name="hNewSort" id="hNewSort" value=""/>
        <br />     
    </td>            
</tr>   
</table>
</form>
</body>
</html>
