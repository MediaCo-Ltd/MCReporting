<%@Language="VBScript"  EnableSessionState="True"%>
<%Option Explicit%> 

 
<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1

Dim Count
Dim Reasons()
Dim ReasonRs
Dim GetReasonSql
Dim Nodata
Dim ReturnData
Dim SentIds

Nodata = False
SentIds = Cstr(Request.Form("xSentIds"))

'If Instr(1,SentIds,"102",1) > 0 Then SentIds = Replace(SentIds,"102","",1,-1,1)
'If Right(SentIds,1) = "," Then SentIds = Left(SentIds,Len(SentIds) -1) 
'If Instr(1,SentIds,",,",1) > 0 Then SentIds = Replace(SentIds,",,",",",1,-1,1)


'If SentType = 4 Then '## All types    
    GetReasonSql = "SELECT DISTINCT Description, GroupId, Id, HasLogs, Usage, SortOrder"
    GetReasonSql = GetReasonSql & " FROM Reasons"
    GetReasonSql = GetReasonSql & " WHERE (GroupId In( 0," & SentIds & "))"
    GetReasonSql = GetReasonSql & " ORDER BY SortOrder"          
'Else
'    GetReasonSql = "SELECT DISTINCT Logs.GroupId, Logs.GroupSelection, Reasons.Description, Reasons.HasLogs, Reasons.Id"
'    GetReasonSql = GetReasonSql & " FROM Logs INNER JOIN"
'    GetReasonSql = GetReasonSql & " Reasons ON Logs.GroupSelection = Reasons.Id"
'    GetReasonSql = GetReasonSql & " WHERE(Reasons.HasLogs = 1) AND (Logs.GroupId = " & SentType & ")"
'    GetReasonSql = GetReasonSql & " ORDER BY Logs.GroupSelection"    
'End If


Set ReasonRs = Server.CreateObject("ADODB.Recordset")
ReasonRs.ActiveConnection = Session("ConnNCReports")
ReasonRs.Source = GetReasonSql
ReasonRs.CursorType = Application("adOpenForwardOnly")
ReasonRs.CursorLocation = Application("adUseClient")
ReasonRs.LockType = Application("adLockReadOnly")
ReasonRs.Open
Set ReasonRs.ActiveConnection = Nothing

If ReasonRs.BOF = True Or ReasonRs.EOF = True Then
    Nodata = True
Else
    Redim Reasons(ReasonRs.RecordCount,1)
    While Not ReasonRs.EOF
        If ReasonRs("Id") < 10 Then
            Reasons(ReasonRs.AbsolutePosition,0) = "0" & ReasonRs("Id")
        Else
            Reasons(ReasonRs.AbsolutePosition,0) = ReasonRs("Id")
        End If
        Reasons(ReasonRs.AbsolutePosition,1) = ReasonRs("Description")
        ReasonRs.MoveNext
    Wend        
End If

ReasonRs.Close
Set ReasonRs = Nothing

If Nodata = False Then

    ReturnData = "<option value='' >Select Issue/Problem</option>" & VbCrLF
    For Count = 1 To Ubound(Reasons)
        ReturnData = ReturnData & ("<option value='" & Cstr(Reasons(Count,0)) & "'>" & Trim(Reasons(Count,1)) & "</option>)") & VbCrLf
    Next
    Erase Reasons

End If

'##ReturnData = SentType
Response.Write ReturnData

%>
