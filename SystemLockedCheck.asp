<%@language="vbscript" codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<% 
On Error Resume Next 

Dim LocalState
Dim StatusRS
		 
Set StatusRS = Server.CreateObject("ADODB.Recordset")
StatusRS.ActiveConnection = Session("ConnStatus")  'strConStatus
StatusRS.Source = "SELECT Locked FROM System"
StatusRS.CursorType = Application("adOpenForwardOnly")
StatusRS.CursorLocation = Application("adUseClient")
StatusRS.LockType = Application("adLockReadOnly")
StatusRS.Open
Set StatusRS.ActiveConnection = Nothing

If Err <> 0 Then
    LocalState = False
Else
    If StatusRS("Locked") = True Then
        StatusRS.Close
        Set StatusRS = Nothing
        LocalState = True
    Else
        StatusRS.Close
        Set StatusRS = Nothing
        LocalState = False
    End If
End If


Err.Clear
Session("Locked") = LocalState

%>