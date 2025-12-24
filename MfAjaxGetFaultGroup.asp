<%@Language="VBScript"  EnableSessionState="True"%>
<%Option Explicit%> 

 
<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1

Dim FaultGroupRs
Dim FaultGroupSql
Dim FaultGroupSqlWhere

Dim Count
Dim FaultGroup()

Dim Nodata
Dim ReturnData
Dim GetGlobal
Dim MachineType
Dim ReturnGroupValue

Nodata = False
ReturnGroupValue = False
MachineType = Cstr(Request.Form("xMg"))

If Left(MachineType,1) = "#" Then
    MachineType = Replace(MachineType,"#","",1,-1,1)
    FaultGroupSqlWhere = " Where (MachineTypeId IN(0," & MachineType & "))"
Else
    FaultGroupSqlWhere = " Where (Id IN(" & MachineType & "))"
End If

FaultGroupSql = "SELECT DISTINCT Id, Description, MachineTypeId, SelectId From FaultGroups"
FaultGroupSql = FaultGroupSql & FaultGroupSqlWhere
FaultGroupSql = FaultGroupSql & " Order By Description"

Set FaultGroupRs = Server.CreateObject("ADODB.Recordset")
FaultGroupRs.ActiveConnection = Session("ConnMachinefaults") 
FaultGroupRs.Source = FaultGroupSql
FaultGroupRs.CursorType = Application("adOpenForwardOnly")
FaultGroupRs.CursorLocation = Application("adUseClient") 
FaultGroupRs.LockType = Application("adLockReadOnly")
FaultGroupRs.Open
Set FaultGroupRs.ActiveConnection = Nothing

If FaultGroupRs.BOF = True Or FaultGroupRs.EOF = True Then
    Nodata = True
Else
    Redim FaultGroup(FaultGroupRs.RecordCount,1)
    While Not FaultGroupRs.EOF
        
        '## Feb 2018 using fault SelectId, scripts in add page cause problems 
        '## If selected something with Id of 12 you couldn't select somethingwith a Id of 1 or vice versa        
        'FaultGroup(FaultGroupRs.AbsolutePosition,0) = FaultGroupRs("Id")
        
        FaultGroup(FaultGroupRs.AbsolutePosition,1) = FaultGroupRs("Description")
        FaultGroup(FaultGroupRs.AbsolutePosition,0) = FaultGroupRs("SelectId")
        FaultGroupRs.MoveNext
    Wend        
End If

FaultGroupRs.Close
Set FaultGroupRs = Nothing

If Nodata = False Then

    ReturnData = "<option value='' >Select&nbsp;Fault&nbsp;Group</option>"
    For Count = 1 To Ubound(FaultGroup)
        'ReturnData = ReturnData & ("<option value='" & Cstr(FaultGroup(Count,0)) & "#" & Cstr(FaultGroup(Count,2)) & "'>" & Trim(FaultGroup(Count,1)) & "</option>)") & VbCrLf
        ReturnData = ReturnData & ("<option value='" & Cstr(FaultGroup(Count,0)) & "'>" & Trim(FaultGroup(Count,1)) & "</option>)") & VbCrLf
        
    Next
    Erase FaultGroup

End If

Response.Write ReturnData

%>
