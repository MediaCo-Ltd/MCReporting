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
Dim SentType

Nodata = False
SentType = Cint(Request.Form("xSentType"))

If SentType = 4 Then '## All types    
    GetReasonSql = "SELECT DISTINCT Description, GroupId, Id, HasLogs, Usage"
    GetReasonSql = GetReasonSql & " FROM Reasons"
    GetReasonSql = GetReasonSql & " WHERE (HasLogs = 1)"
    GetReasonSql = GetReasonSql & " ORDER BY Usage, Description"          
Else
    GetReasonSql = "SELECT DISTINCT Logs.GroupId, Logs.GroupSelection, Reasons.Description, Reasons.HasLogs, Reasons.Id"
    GetReasonSql = GetReasonSql & " FROM Logs INNER JOIN"
    GetReasonSql = GetReasonSql & " Reasons ON Logs.GroupSelection = Reasons.Id"
    GetReasonSql = GetReasonSql & " WHERE(Reasons.HasLogs = 1) AND (Logs.GroupId = " & SentType & ")"
    GetReasonSql = GetReasonSql & " ORDER BY Logs.GroupSelection"    
End If


Set ReasonRs = Server.CreateObject("ADODB.Recordset")
ReasonRs.ActiveConnection = Session("ConnHSReports")
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
        Reasons(ReasonRs.AbsolutePosition,0) = ReasonRs("Id")
        Reasons(ReasonRs.AbsolutePosition,1) = ReasonRs("Description")
        ReasonRs.MoveNext
    Wend        
End If

ReasonRs.Close
Set ReasonRs = Nothing

If Nodata = False Then

    ReturnData = "<option value='' >Select&nbsp;Reason</option>"
    For Count = 1 To Ubound(Reasons)
        ReturnData = ReturnData & ("<option value='" & Cstr(Reasons(Count,0)) & "'>" & Trim(Reasons(Count,1)) & "</option>)") & VbCrLf
    Next
    Erase Reasons

End If

Response.Write ReturnData

%>
