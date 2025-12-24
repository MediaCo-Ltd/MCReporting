<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 
<%Option Explicit%>


<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<%

If Session("UserName") = "" Then Response.Redirect ("Login.asp")

'################################################ Get  Data 

Dim LogRs
Dim LogSql
Dim RecordToEmail
Dim EmailGroup
Dim EmailSeverity
Dim EmailNotes
Dim EmailResolved

Dim RedirectUrl
RedirectUrl = "HsSelectOption.asp"

RecordToEmail = Clng(Session("EmailId"))
Session("EmailId") = ""

LogSql = "SELECT  Logs.Id, Logs.UserId, Logs.GroupId, Logs.GroupSelection, Logs.Notes, Logs.CreatedDate, Logs.CreatedDateSerial,"
LogSql = LogSql & " Logs.Severity, Logs.SelectedDate, Logs.SelectedDateSerial, Logs.Resolved, Logs.CreatedByName,"
LogSql = LogSql & " Reasons.Description"
LogSql = LogSql & " From Logs INNER JOIN Reasons ON Logs.GroupSelection = Reasons.Id"
'LogSql = LogSql & " INNER JOIN Reasons ON Logs.GroupSelection = Reasons.Id"
LogSql = LogSql & " Where (Logs.Id = " & RecordToEmail & ")"

Set LogRs = Server.CreateObject("ADODB.Recordset")
LogRs.ActiveConnection = Session("ConnHSReports")
LogRs.Source = LogSql  
LogRs.CursorType = Application("adOpenForwardOnly")
LogRs.CursorLocation = Application("adUseClient") 
LogRs.LockType = Application("adLockReadOnly")
LogRs.Open
Set LogRs.ActiveConnection = Nothing

If LogRs.BOF = True Or LogRs.EOF = True Then
    LogRs.Close
    Set LogRs = Nothing
    Response.Redirect (RedirectUrl)
Else
    If LogRs("GroupId") = 1 Then
        EmailGroup = "Accident"
    ElseIf LogRs("GroupId") = 2 Then
        EmailGroup = "Incident"
    Else
        EmailGroup = "Unsafe Condition or Damage"
    End If
    
    If LogRs("Severity") = 1 Then
        EmailSeverity = "Minor"
    ElseIf LogRs("Severity") = 2 Then
        EmailSeverity = "Medium"
    Else
        EmailSeverity = "Critical"
    End If
    
    If LogRs("Resolved") = Cbool(True) Then
        EmailResolved = "Yes"
    Else
        EmailResolved = "No"
    End If
    
    If LogRs("Notes") <> "" Then EmailNotes = Trim(LogRs("Notes"))
   
End If

'############################################ Set up email

Dim WhoTo
Dim WhoToName
Dim WhoToCC
Dim WhoToNameCC
Dim Emsg
Dim MsgHtml
Dim ContentId1

Set Emsg = Server.CreateOBject( "JMail.Message" )

Emsg.Logging = true
Emsg.Silent = true
Emsg.From = "no-reply@mediaco.co.uk"
Emsg.FromName = "HS Reporting"
Emsg.Subject = "HS Reporting New " & EmailGroup & " Record #" & RecordToEmail

If Session("Smtp") = "mx496502.smtp-engine.com" Then 	    
    '## For testing at home
    ContentId1 = Emsg.AddAttachment(Session("RootPath") & "\Images\Plus-icon-email.png",true,"image/png")
ElseIf Session("Smtp") = "192.168.20.50" Or Session("Smtp") = "192.168.20.37" Then          
    '## Work my Pc 
    ContentId1 = Emsg.AddAttachment(Session("RootPath") & "\Images\Plus-icon-email.png",true,"image/png")   
End If

Emsg.HTMLBody = CreateMsg

'### Send email 
If Session("PC-Name") = "Server" Then    
    SendEmail    
Else
    '## Send mail from home or work,except for some testing I don't need it to
    '## SendEmail 
End If

LogRs.Close
Set LogRs = Nothing

Response.Redirect (RedirectUrl)

'###################################################################

Sub SendEmail()

Dim NotifyRs

If Session("Smtp") = "mx496502.smtp-engine.com" Then

       
    '## For testing at home & if go to smarthost
    Emsg.MailServerUserName = "warren.morris@mediaco.co.uk"
    Emsg.MailServerPassWord = "Q6)9l7Sit6dAV*"
    
    If Session("PC-Name") = "Work" Then        
        '## For testing at work
        Emsg.AddRecipient "alan@mediaco.co.uk", "Alan Holgate" 
    ElseIf Session("PC-Name") = "Server" Then
        If Session("UserId") = 1 Or Session("UserId") = 4 Then
            Emsg.AddRecipient "alan@mediaco.co.uk", "Alan Holgate" 
        Else
            '## Send email to all on email list            
            
            Set NotifyRs = Server.CreateObject("ADODB.Recordset")
            
            NotifyRs.ActiveConnection = Session("ConnMcLogon")
            NotifyRs.Source = "SELECT EmailName, EmailAddress, EmailActive FROM Email Where (EmailActive = 1) AND (EmailHS = 1) Order By EmailName"
            NotifyRs.CursorType = Application("adOpenForwardOnly")
            NotifyRs.CursorLocation = Application("adUseClient")
            NotifyRs.LockType = Application("adLockReadOnly")
            NotifyRs.Open
            
            While Not NotifyRs.EOF
                Emsg.AddRecipient Trim(NotifyRs("EmailAddress")), Trim(NotifyRs("EmailName"))                    
                NotifyRs.MoveNext
            Wend
            
            NotifyRs.Close
            Set NotifyRs = Nothing            
        End If
    Else
        Emsg.AddRecipient "alan.holgate@yahoo.co.uk", "Alan Holgate"
    End If 
    
ElseIf Session("Smtp") = "192.168.20.50"  Or Session("Smtp") = "192.168.20.37" Then
    If Session("PC-Name") = "Work" Then        
        '## For testing at work
        Emsg.AddRecipient "alan@mediaco.co.uk", "Alan Holgate" 
    Else
        If Session("UserId") = 1 Or Session("UserId") = 4 Then
            Emsg.AddRecipient "alan@mediaco.co.uk", "Alan Holgate" 
        Else
            '## Send email to all on email list            
            
            Set NotifyRs = Server.CreateObject("ADODB.Recordset")
            
            NotifyRs.ActiveConnection = Session("ConnMcLogon")
            NotifyRs.Source = "SELECT EmailName, EmailAddress, EmailActive FROM Email Where (EmailActive = 1) AND (EmailHS = 1) Order By EmailName"
            NotifyRs.CursorType = Application("adOpenForwardOnly")
            NotifyRs.CursorLocation = Application("adUseClient")
            NotifyRs.LockType = Application("adLockReadOnly")
            NotifyRs.Open
            
            While Not NotifyRs.EOF
                Emsg.AddRecipient Trim(NotifyRs("EmailAddress")), Trim(NotifyRs("EmailName"))                    
                NotifyRs.MoveNext
            Wend
            
            NotifyRs.Close
            Set NotifyRs = Nothing
            
        End If
    End If
End If  

'## Emsg.AddRecipientCC "alan@mediaco.co.uk", "Alan Holgate" 
'Emsg.AddRecipient "alan@mediaco.co.uk", "Alan Holgate" 

Emsg.Logging = true

If Not Emsg.Send(Session("Smtp")) Then
    Session("SystemError") = "HS Email send fail, " & Emsg.ErrorMessage & " Record " & RecordToEmail & ". User = " & Session("UserName") & ". SessionId = " & Session.SessionID
    RedirectUrl = "SystemError.asp"
    'Response.Redirect "SystemError.asp"
Else
    UpdateEmailStatus    
End If

Emsg.Close
Set Emsg = nothing

End Sub

'############################################################ Update email status

Sub UpdateEmailStatus

    Dim EmailRs

    Set EmailRs = Server.CreateObject("ADODB.Recordset")
    EmailRs.ActiveConnection = Session("ConnHSReports")
    EmailRs.Source = "SELECT Id, Emailed FROM Logs WHERE (id = " & RecordToEmail & ")"
    EmailRs.CursorType = Application("adOpenStatic")
    EmailRs.CursorLocation = Application("adUseClient")
    EmailRs.LockType = Application("adLockOptimistic")
    EmailRs.Open

    EmailRs("Emailed") = 1

    EmailRs.Update

    Set EmailRs.ActiveConnection = Nothing
    EmailRs.Close
    Set EmailRs = Nothing

    Err.Clear

End sub

'#####################################

Function CreateMsg()

'***********************
'   Create Html Document 
'***********************

MsgHtml = ""

MsgHtml =  "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>" & VbCrLf
MsgHtml = MsgHtml & "<html xmlns='http://www.w3.org/1999/xhtml' xml:lang='en' lang='en'>" & VbCrLf

MsgHtml = MsgHtml & "<head>" & VbCrLf
MsgHtml = MsgHtml & "<title>HS Report</title>" & VbCrLf
MsgHtml = MsgHtml & "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />" & VbCrLf
MsgHtml = MsgHtml & "<style type='text/css'>" & VbCrLf
MsgHtml = MsgHtml & "body" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "font-family:  'Verdana', 'Arial', 'Geneva',  'Lucida', 'San-Serif', 'Lucida Sans', 'Lucida Console', 'MS Sans Serif';" & VbCrLf
MsgHtml = MsgHtml & "font-size: 69%;" & VbCrLf
MsgHtml = MsgHtml & "font-style: normal;" & VbCrLf
MsgHtml = MsgHtml & "font-weight: normal;" & VbCrLf
MsgHtml = MsgHtml & "font-variant:  normal;" & VbCrLf
MsgHtml = MsgHtml & "text-transform:  none;" & VbCrLf
MsgHtml = MsgHtml & "text-decoration:  none;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".StdLabel" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "padding-top: 0px;" & VbCrLf
MsgHtml = MsgHtml & "padding-left: 10px;" & VbCrLf
MsgHtml = MsgHtml & "padding-bottom: 0px;" & VbCrLf
MsgHtml = MsgHtml & "padding-right: 0px;" & VbCrLf
MsgHtml = MsgHtml & "vertical-align: bottom;" & VbCrLf
MsgHtml = MsgHtml & "color: #0069AA;" & VbCrLf
MsgHtml = MsgHtml & "font-weight: bold;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".HeaderText" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "color: #000000;" & VbCrLf
MsgHtml = MsgHtml & "padding-top: 0px;" & VbCrLf
MsgHtml = MsgHtml & "padding-left: 0px;" & VbCrLf
MsgHtml = MsgHtml & "padding-bottom: 3px;" & VbCrLf
MsgHtml = MsgHtml & "padding-right: 0px;" & VbCrLf
MsgHtml = MsgHtml & "font-weight: bold;" & VbCrLf
MsgHtml = MsgHtml & "font-size:medium;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".StdText" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "color: #000000;" & VbCrLf
MsgHtml = MsgHtml & "font-weight: bold;" & VbCrLf
MsgHtml = MsgHtml & "padding-top: 0px;" & VbCrLf
MsgHtml = MsgHtml & "padding-left: 10px;" & VbCrLf
MsgHtml = MsgHtml & "padding-bottom: 0px;" & VbCrLf
MsgHtml = MsgHtml & "padding-right: 0px;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".NotesText" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "color: #000000;" & VbCrLf
MsgHtml = MsgHtml & "padding-top: 5px;" & VbCrLf
MsgHtml = MsgHtml & "padding-left: 10px;" & VbCrLf
MsgHtml = MsgHtml & "vertical-align: top;" & VbCrLf
MsgHtml = MsgHtml & "font-size: 11px;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & "</style>" & VbCrLf
MsgHtml = MsgHtml & "</head>" & VbCrLf
MsgHtml = MsgHtml & "<body>" & VbCrLf

MsgHtml = MsgHtml & "<table width='100%' >" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<td align='left' valign='bottom' width='25%' >" & VbCrLf
MsgHtml = MsgHtml & "<img align='left' alt='HS logo' width='30px' src=""cid:" & ContentId1 & """ /> " & VbCrLf  ' width ='90px'
MsgHtml = MsgHtml & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td class='HeaderText' align='center'>HS Reporting New " & EmailGroup & " Record #" & RecordToEmail & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td align='right' valign='bottom' width='25%' >&nbsp;</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "</table>" & VbCrLf


MsgHtml = MsgHtml & "<br />" & VbCrLf
'MsgHtml = MsgHtml & "<br />" & VbCrLf

MsgHtml = MsgHtml & "<table style='width: 100%;' >" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<td valign='top' style='border-style: solid none none none; border-top-width: 2px; border-top-color: #19ABDE; height: 12px;' colspan='2'></td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "</table>" & VbCrLf


MsgHtml = MsgHtml & "<br />" & VbCrLf

MsgHtml = MsgHtml & "<table style='width: 90%;' align='center' >" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='20%'>Incident&nbsp;Date</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='20%'>Created&nbsp;By</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='20%'>Reason</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='20%'>Severity</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='20%'>Resolved</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf

MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & Left(LogRs("SelectedDate"),16) & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & LogRs("CreatedByName") & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & LogRs("Description") & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & EmailSeverity & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & EmailResolved & "</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & " <td height='3px' colspan='5'></td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf

MsgHtml = MsgHtml & "</table>" & VbCrLf
MsgHtml = MsgHtml & "<table style='width: 90%;' align='center'  >" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' >Comments</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "<tr >" & VbCrLf
MsgHtml = MsgHtml & "<td class='NotesText'>" & VbCrLf   
MsgHtml = MsgHtml & EmailNotes
MsgHtml = MsgHtml & "</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "</table>" & VbCrLf

MsgHtml = MsgHtml & "<br />" & VbCrLf
MsgHtml = MsgHtml & "<br />" & VbCrLf
MsgHtml = MsgHtml & "<br />" & VbCrLf

'MsgHtml = MsgHtml & "<table style='width: 98%; padding-right: 10px; padding-left: 10px;'>" & VbCrLf
'MsgHtml = MsgHtml & "<tr>" & VbCrLf
'MsgHtml = MsgHtml & "<td height='50'>" & VbCrLf
'MsgHtml = MsgHtml & "<p align='center'><font size='2.5'> MediaCo Production Ltd. Churchill Point, Churchill Way<br />" 
'MsgHtml = MsgHtml & "Trafford Park, Manchester M17 1BS <br />Tel: (+44)161 875 2020 Fax: (+44)161 873 7740</font></p>" & VbCrLf
'MsgHtml = MsgHtml & "</td>" & VbCrLf
'MsgHtml = MsgHtml & "</tr>" & VbCrLf
'MsgHtml = MsgHtml & "</table>" & VbCrLf

MsgHtml = MsgHtml & "</body>" & VbCrLf
MsgHtml = MsgHtml & "</html>" & VbCrLf


CreateMsg = MsgHtml

End Function
%>

