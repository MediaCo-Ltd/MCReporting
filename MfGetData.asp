<% 
Dim errMsg
Dim LogSql
Dim MachineSql
Dim MachineRs
Dim LogRs
Dim MyStartDate
Dim MyEndDate
Dim strExtra
Dim strUserExtra

errMsg = ""
strExtra = ""
strUserExtra = ""

'Machine.Id AS Mid, Machine.MachineTypeId AS Mtid,  Machine.MachineType,
LogSql = "SELECT  Distinct Logs.Id, Logs.UserId, Logs.MachineId, Logs.MachineTypeId, Logs.Status, Logs.ErrorNotes, Logs.LogDate, Logs.LogDateSerial,"
LogSql = LogSql & " Logs.RepairNotes, Logs.RepairDate, Logs.RepairDateSerial, Logs.Cost, Logs.RecurringFault, Logs.FaultRepaired,"
LogSql = LogSql & " Logs.ErrorDescription, Logs.FaultGroups, Machine.MachineName, Machine.Active, Logs.CreatedByName, Logs.Dormant, Logs.HasImage"
LogSql = LogSql & " FROM Logs INNER JOIN Machine ON Logs.MachineId = Machine.Id"

MachineSql = "SELECT Distinct  Machine.Id, Machine.MachineName, Machine.MachineTypeId"  'Logs.FaultRepaired,
MachineSql = MachineSql & " From Logs Inner join Machine On Logs.MachineId = Machine.Id"

'############################## To split machines into groups 

Sub OpenMachine(StatusFilter,Uid)    

    strExtra = ""
    strUserExtra = ""
    
    If Uid <> "0" Then strUserExtra = " And (Logs.UserId = " & Clng(Uid) & ")"
    
    If StatusFilter = 0 Then
        strExtra = " Where (FaultRepaired = 0) AND (Logs.Dormant = 0) " & strUserExtra & " Order By MachineTypeId, MachineName"    
    ElseIf StatusFilter = 1 Then
        strExtra = " Where (FaultRepaired = 1) AND (Logs.Dormant = 0) " & strUserExtra & " Order By MachineTypeId, MachineName"
    Else
        strExtra = " Where (Dormant = 0) " & strUserExtra & " Order By MachineTypeId, MachineName"
    End If
        
    Set MachineRs = Server.CreateObject("ADODB.Recordset")

	MachineRs.ActiveConnection = Session("ConnMachinefaults") 
	MachineRs.Source = MachineSql & strExtra 
	MachineRs.CursorType = Application("adOpenForwardOnly")
	MachineRs.CursorLocation = Application("adUseClient") 
	MachineRs.LockType = Application("adLockReadOnly")
	MachineRs.Open
	Set MachineRs.ActiveConnection = Nothing

End Sub

'############################## Log data

Sub OpenLogs(MachineId,StatusFilter,Uid)

    strExtra = ""    
    strUserExtra = ""
    
    If Uid <> "0" Then strUserExtra = " And (Logs.UserId = " & Clng(Uid) & ")"
    
    If StatusFilter = 0 Then
        strExtra = " Where (FaultRepaired = 0) And (MachineId = " & MachineId & ") AND (Logs.Dormant = 0) " & strUserExtra & " Order By Status"    
    ElseIf StatusFilter = 1 Then
        strExtra = " Where (FaultRepaired = 1) And (MachineId = " & MachineId & ") AND (Logs.Dormant = 0) " & strUserExtra & " Order By Status"
    Else
        strExtra = " Where (MachineId = " & Clng(MachineId) & ") AND (Logs.Dormant = 0) " & strUserExtra & " Order By Status"
    End If
        
    Set LogRs = Server.CreateObject("ADODB.Recordset")

	LogRs.ActiveConnection = Session("ConnMachinefaults") 
	LogRs.Source = LogSql & strExtra 
	LogRs.CursorType = Application("adOpenForwardOnly")
	LogRs.CursorLocation = Application("adUseClient") 
	LogRs.LockType = Application("adLockReadOnly")
	LogRs.Open
	Set LogRs.ActiveConnection = Nothing


End Sub

Private Function FaultGroupDescription (GrpId)

    Dim GrpSql
    Dim GrpRs
    Dim GrpDesc
    GrpSql = "SELECT Id, Description FROM FaultGroups Where (Id IN(" & GrpId & ")) AND (Active = 1) "

    Set GrpRs = Server.CreateObject("ADODB.Recordset")
    GrpRs.ActiveConnection = Session("ConnMachinefaults")
    GrpRs.Source = GrpSql
    GrpRs.CursorType = Application("adOpenForwardOnly")
    GrpRs.CursorLocation = Application("adUseClient")
    GrpRs.LockType = Application("adLockReadOnly")
    GrpRs.Open

    If GrpRs.BOF= True Or GrpRs.EOF = True Then
        GrpDesc = ""
    Else
        While Not GrpRs.EOF    
            If GrpDesc = "" Then
                GrpDesc = GrpRs("Description")
            Else
                GrpDesc = GrpDesc & ", " & GrpRs("Description")
            End If
            GrpRs.MoveNext
        Wend
    End If

    GrpRs.Close
    Set GrpRs = Nothing

    FaultGroupDescription = GrpDesc

End Function

%>