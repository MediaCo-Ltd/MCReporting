<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 
<%Option Explicit%>


<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<%

If Session("UserName") = "" Then Response.Redirect ("Login.asp")

'################################################ Get  Data 

Dim LogRs
Dim LogSql
Dim RecordToEmail
Dim EmailMachine
Dim EmailSeverity
Dim EmailNotes
Dim EmailRepaired
Dim EmailRecurringFault
Dim EmailDescription
Dim EmailFaultGroup

Dim RedirectUrl
RedirectUrl = "MfSelectOption.asp"

RecordToEmail = Clng(Session("EmailId"))
Session("EmailId") = ""

LogSql = "SELECT  Logs.Id, Logs.UserId, Logs.MachineId, Logs.MachineTypeId, Logs.Status, Logs.ErrorNotes, Logs.LogDate, Logs.LogDateSerial,"
LogSql = LogSql & " Logs.RepairNotes, Logs.RepairDate, Logs.RepairDateSerial, Logs.Cost, Logs.RecurringFault, Logs.FaultRepaired,"
LogSql = LogSql & " Logs.ErrorDescription, Logs.FaultGroups, Machine.MachineName, Machine.MachineType, Machine.Active, Logs.CreatedByName,"
LogSql = LogSql & " Logs.HasImage FROM Logs INNER JOIN Machine ON Logs.MachineId = Machine.Id"
'LogSql = LogSql & " INNER JOIN Users ON Logs.UserId = Users.Id"
LogSql = LogSql & " Where (Logs.Id = " & RecordToEmail & ")"

Set LogRs = Server.CreateObject("ADODB.Recordset")
LogRs.ActiveConnection = Session("ConnMachinefaults")
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
    
    EmailMachine = LogRs("MachineName")
    EmailDescription = LogRs("ErrorDescription")
    '## need to get data   
    EmailFaultGroup = GetFaults (LogRs("FaultGroups"))                  'Replace(LogRs("FaultGroups"),",","<br />",1,-1,1)
    
    If LogRs("Status") = 1 Then   
        EmailSeverity = "Minor"
    ElseIf LogRs("Status") = 2 Then
        EmailSeverity = "Medium"
    Else
        EmailSeverity = "Critical"
    End If
    
    If LogRs("FaultRepaired") = Cbool(True) Then
        EmailRepaired = "Yes"
    Else
        EmailRepaired = "No"
    End If
    
    If LogRs("RecurringFault") = Cbool(True) Then
        EmailRecurringFault = "Yes"
    Else
        EmailRecurringFault = "No"
    End If
    
    If LogRs("ErrorNotes") <> "" Then EmailNotes = Trim(LogRs("ErrorNotes"))
   
End If

Dim UserIpAddress 
Dim Url
UserIpAddress = Request.ServerVariables("REMOTE_ADDR")
If Left(UserIpAddress,10) = "192.168.20" Then
    Url = "http://192.168.20.3"
Else
    Url = "http://http://82.71.163.186"
End If

Dim Imagelink
Imagelink = ""
If LogRs("HasImage") = Cbool(True) Then
    Imagelink = Url & "/MCReporting/MfShowRemoteImages.asp?id=" & RecordToEmail & "&vr=1"
End If



'http://127.0.0.1/MCReporting/MfShowRemoteImages.asp?Id=38&vr=1

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
Emsg.FromName = "Machine Faults"
Emsg.Subject = "Fault on " & EmailMachine '&  Record #" & RecordToEmail


'## Screwdriver image is rubbish
If Session("Smtp") = "mx496502.smtp-engine.com" Then 	    
    '## For testing at home  
    ContentId1 = Emsg.AddAttachment(Session("RootPath") & "\Images\W-S48.png",true,"image/png")
ElseIf Session("Smtp") = "192.168.20.50"  Or Session("Smtp") = "192.168.20.37" Then          
    '## Work my Pc 
    ContentId1 = Emsg.AddAttachment(Session("RootPath") & "\Images\W-S48.png",true,"image/png") 
End If

Emsg.HTMLBody = CreateMsg

'### Send email 
If Session("Smtp") = "mx496502.smtp-engine.com" Then 
    '## Mail will go out from home under these settings
    '## But except for some testing I don't need it to
    SendEmail    
Else
    SendEmail 
End If

LogRs.Close
Set LogRs = Nothing

Response.Redirect (RedirectUrl)

'###################################################################

Function GetFaults(FaultList)

    Dim FaultsRs
    Dim LocalResult

    Set FaultsRs = Server.CreateObject("ADODB.Recordset")
    FaultsRs.ActiveConnection = Session("ConnMachinefaults")
    FaultsRs.Source = "SELECT Id, Description FROM FaultGroups WHERE (Id IN (" & FaultList & "))"  
    FaultsRs.CursorType = Application("adOpenForwardOnly")
    FaultsRs.CursorLocation = Application("adUseClient") 
    FaultsRs.LockType = Application("adLockReadOnly")
    FaultsRs.Open

    While Not FaultsRs.EOF
        If LocalResult = "" Then
            LocalResult = FaultsRs("Description") 
        Else
            LocalResult = LocalResult & "<br />" & FaultsRs("Description")
        End If
        FaultsRs.MoveNext    
    Wend
    
    FaultsRs.Close
    Set FaultsRs = Nothing

    GetFaults = LocalResult
    
End Function    

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
            '## All by groups so no need for this 
            
            If EmailSeverity = "Critical" Then
                Emsg.AddRecipient "MFCGE@mediaco.co.uk", "MFCGE" 
            Else
                Emsg.AddRecipient "MFGE@mediaco.co.uk", "MFGE"  
            End If
            ' MFCGE else MFGE   
            
           ' Set NotifyRs = Server.CreateObject("ADODB.Recordset")
            
           ' NotifyRs.ActiveConnection = Session("ConnMcLogon")
           ' NotifyRs.Source = "SELECT EmailName, EmailAddress, EmailActive FROM Email Where (EmailActive = 1) AND (EmailMF = 1) Order By EmailName"
           ' NotifyRs.CursorType = Application("adOpenForwardOnly")
           ' NotifyRs.CursorLocation = Application("adUseClient")
           ' NotifyRs.LockType = Application("adLockReadOnly")
           ' NotifyRs.Open
            
           ' While Not NotifyRs.EOF
           '     Emsg.AddRecipient Trim(NotifyRs("EmailAddress")), Trim(NotifyRs("EmailName"))                    
           '     NotifyRs.MoveNext
           ' Wend
            
           ' NotifyRs.Close
           ' Set NotifyRs = Nothing            
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
            '## All by groups so no need for this
            
            If EmailSeverity = "Critical" Then
                Emsg.AddRecipient "MFCGE@mediaco.co.uk", "MFCGE" 
            Else
                Emsg.AddRecipient "MFGE@mediaco.co.uk", "MFGE"  
            End If
          
          
          
          
          '  Set NotifyRs = Server.CreateObject("ADODB.Recordset")
            
          '  NotifyRs.ActiveConnection = Session("ConnMcLogon")
         '   NotifyRs.Source = "SELECT EmailName, EmailAddress, EmailActive FROM Email Where (EmailActive = 1) AND (EmailMF = 1) Order By EmailName"
         '   NotifyRs.CursorType = Application("adOpenForwardOnly")
          '  NotifyRs.CursorLocation = Application("adUseClient")
          '  NotifyRs.LockType = Application("adLockReadOnly")
          '  NotifyRs.Open
            
          '  While Not NotifyRs.EOF
          '      Emsg.AddRecipient Trim(NotifyRs("EmailAddress")), Trim(NotifyRs("EmailName"))                    
          '      NotifyRs.MoveNext
          '  Wend
            
          '  NotifyRs.Close
          '  Set NotifyRs = Nothing
            
        End If
    End If
End If


'EmailSeverity = "Critical" then MFCGE else MFGE


Emsg.Logging = true

'## Emsg.AddRecipientCC "alan@mediaco.co.uk", "Alan Holgate" 

'## Emsg.AddRecipient "alan@mediaco.co.uk", "Alan Holgate"


If Not Emsg.Send(Session("Smtp")) Then
    Session("SystemError") = "MF Email send fail, " & Emsg.ErrorMessage & " Record " & RecordToEmail & ". User = " & Session("UserName") & ". SessionId = " & Session.SessionID
    RedirectUrl = "SystemError.asp"
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
    EmailRs.ActiveConnection = Session("ConnMachinefaults")
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
MsgHtml = MsgHtml & "<title>Machine Faults</title>" & VbCrLf
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
MsgHtml = MsgHtml & "<img align='left' alt='logo' width='50px' src=""cid:" & ContentId1 & """ /> " & VbCrLf  
MsgHtml = MsgHtml & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td class='HeaderText' align='center'>New Fault log for " & EmailMachine & " Record #" & RecordToEmail & "</td>" & VbCrLf
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
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='15%'>Date</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='20%'>Created&nbsp;By</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='26%'>Machine</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='15%'>Severity</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='12%'>Repaired</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='12%'>Recurring</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & LogRs("LogDate") & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & LogRs("CreatedByName") & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & EmailMachine & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & EmailSeverity & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & EmailRepaired & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdText' align='left'>" & EmailRecurringFault & "</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & " <td height='3px' colspan='6'></td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "</table>" & VbCrLf

MsgHtml = MsgHtml & "<br />" & VbCrLf

MsgHtml = MsgHtml & "<table style='width: 90%;' align='center'  >" & VbCrLf
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' width='50%'>FaultsComments</td>" & VbCrLf
MsgHtml = MsgHtml & "<td  class='StdLabel' align='left' >Comments</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "<tr >" & VbCrLf
MsgHtml = MsgHtml & "<td class='NotesText'>" & VbCrLf   
MsgHtml = MsgHtml & EmailFaultGroup
MsgHtml = MsgHtml & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td class='NotesText'>" & VbCrLf
MsgHtml = MsgHtml & EmailNotes
MsgHtml = MsgHtml & "</td>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "</table>" & VbCrLf

'MsgHtml = MsgHtml & "<br />" & VbCrLf
'MsgHtml = MsgHtml & "<br />" & VbCrLf

If Imagelink <> "" Then
    MsgHtml = MsgHtml & "<br />" & VbCrLf
    MsgHtml = MsgHtml & "<br />" & VbCrLf
    MsgHtml = MsgHtml & "<table style='width: 98%; padding-right: 10px; padding-left: 10px;'>" & VbCrLf
    MsgHtml = MsgHtml & "<tr>" & VbCrLf
    MsgHtml = MsgHtml & "<td class='StdLabel' align='left' height='50'>" & VbCrLf
    MsgHtml = MsgHtml & "<a  href='" & Imagelink & " style='font-size:14px; color: " & NewBlue & "'>View Images</a>"  
    MsgHtml = MsgHtml & "</td>" & VbCrLf
    MsgHtml = MsgHtml & "</tr>" & VbCrLf
    MsgHtml = MsgHtml & "</table>" & VbCrLf    
End If

MsgHtml = MsgHtml & "</body>" & VbCrLf
MsgHtml = MsgHtml & "</html>" & VbCrLf

CreateMsg = MsgHtml

End Function
%>

