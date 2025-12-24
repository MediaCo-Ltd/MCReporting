<%@Language="VBScript"  EnableSessionState="True"%>
<%Option Explicit%> 

 
<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1

Dim Count
Dim Users()
Dim UserRs
Dim GetUserSql
Dim Nodata
Dim ReturnData
Dim SentType
Dim SentReason

Nodata = False
SentType = Cint(Request.Form("xType"))
SentReason = Cint(Request.Form("xReason"))

If SentType = 4 Then '## All Users    
    If SentReason = 0 Then
        GetUserSql = "SELECT DISTINCT UserId, CreatedByName"
        GetUserSql = GetUserSql & " FROM Logs"
        GetUserSql = GetUserSql & " ORDER BY CreatedByName"
    Else
        GetUserSql = "SELECT DISTINCT UserId, GroupSelection, Logs.GroupId, CreatedByName"  
        GetUserSql = GetUserSql & " FROM Logs"
        GetUserSql = GetUserSql & " WHERE (GroupSelection = " & SentReason & ")"
        GetUserSql = GetUserSql & " ORDER BY CreatedByName"
    End If      
Else
    If SentReason = 0 Then
        GetUserSql = "SELECT DISTINCT UserId, CreatedByName" 
        GetUserSql = GetUserSql & " FROM Logs"
        GetUserSql = GetUserSql & " WHERE (GroupId = " & SentType & ")"
        GetUserSql = GetUserSql & " ORDER BY CreatedByName"
    Else    
        GetUserSql = "SELECT DISTINCT UserId, GroupSelection, GroupId, CreatedByName"   
        GetUserSql = GetUserSql & " FROM Logs"
        GetUserSql = GetUserSql & " WHERE (GroupId = " & SentType & ") AND (GroupSelection = " & SentReason & ")"
        GetUserSql = GetUserSql & " ORDER BY CreatedByName"
    End If   
End If

Set UserRs = Server.CreateObject("ADODB.Recordset")
UserRs.ActiveConnection = Session("ConnHSReports")
UserRs.Source = GetUserSql
UserRs.CursorType = Application("adOpenForwardOnly")
UserRs.CursorLocation = Application("adUseClient")
UserRs.LockType = Application("adLockReadOnly")
UserRs.Open
Set UserRs.ActiveConnection = Nothing

If UserRs.BOF = True Or UserRs.EOF = True Then
    Nodata = True
Else
    Redim Users(UserRs.RecordCount,1)
    While Not UserRs.EOF
        Users(UserRs.AbsolutePosition,0) = UserRs("UserId")
        Users(UserRs.AbsolutePosition,1) = UserRs("CreatedByName")
        UserRs.MoveNext
    Wend        
End If

UserRs.Close
Set UserRs = Nothing

If Nodata = False Then

    ReturnData = "<option value='' >Select&nbsp;User</option>"
    For Count = 1 To Ubound(Users)
        ReturnData = ReturnData & ("<option value='" & Cstr(Users(Count,0)) & "'>" & Trim(Users(Count,1)) & "</option>)") & VbCrLf
    Next
    Erase Users
Else
    ReturnData = ""
End If

Response.Write ReturnData

%>
