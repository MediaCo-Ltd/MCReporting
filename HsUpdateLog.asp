<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp?uf=hs")

Dim Error
Dim UserId
Dim UpdateOk
Dim Count
Dim Notes
Dim NewNotes
Dim SaveNotes

Dim AddRs
Dim AddSql
Dim UpdateRs
Dim UpdateSql
Dim RedirectUrl
Dim LocalTimeStr
Dim EditRecordId
Dim NewRecordId
Dim UpdateReasonId

RedirectUrl = "HsSelectOption.asp"

If Request.Form("chkUpdate") = "on" Then
    '## Testing don't save data, will only be on if at home or on my pc at work
Else

    UpdateReasonId = Cint(Request.Form("hReasonId"))
    
'############################## Add New Log ##############################
    If Request.Form("frmName") = "frmAddLog" Then  

    '############################## Create log 

        AddSql = "SELECT * FROM Logs" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnHSReports")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open

        AddRs.AddNew

        AddRs("GroupId") = Cint(Request.Form("hGroupId"))
        AddRs("GroupSelection") = Cint(Request.Form("hReasonId"))
        AddRs("Severity") = Cint(Request.Form("hSeverity"))  
        
        If Request.Form("hRiddor") > 0 Then
            AddRs("Riddor") = Cbool(Request.Form("hRiddor"))
            AddRs("RiddorLevel") = Cint(Request.Form("hRiddor"))
            AddRs("RiddorDays") = Cint(Request.Form("hRiddorDays"))
        End If
        
        If Request.Form("hRiddor") > 0 Then 
            AddRs("RiddorDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        Else
            AddRs("RiddorDateSerial") = 0
        End If
       
        If Request.Form("hLocation") <> "" Then
            If Cint(Request.Form("hLocation")) > 0 Then AddRs("LocationId") = Cint(Request.Form("hLocation"))
        End If
        
        If Request.Form("txtDetails") <> "" Then        
            '## need to do formatting so vbcrlf = <br />            
            Notes = Trim(Cstr(Request.Form("txtDetails")))
            Notes = Replace(Notes,VbCrLf,"<br />",1,-1,1)
            '## If triple line feeds cut one out
            '## They seem to like having lots of gaps, making the notes to long
            '## Triple = 1 for return of 1st line then 2 gaps before next line
            Notes = Replace(Notes,"<br /><br /><br />","<br /><br />",1,-1,1)
            Notes = Replace(Notes,"&","&amp;",1,-1,1)
            AddRs("Notes") = Trim(Notes)
        End If
        
        AddRs("CreatedDate") = Date
        AddRs("CreatedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))        
        
        
        
        If Request.Form("txtDate") <> "" Then
          '  If Request.Form("txtDate") > Now Then
          '      LocalTimeStr = Time()
          '      AddRs("SelectedDate") = Now
          '      AddRs("SelectedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
          '  Else 
                LocalTimeStr = Right(Request.Form("txtDate"),5) & ":00"
                AddRs("SelectedDate") = Request.Form("txtDate")
                AddRs("SelectedDateSerial") = CDbl(DateSerial(Year(Request.Form("txtDate")),Month(Request.Form("txtDate")),Day(Request.Form("txtDate"))))
          '  End If
        Else    
            LocalTimeStr = Time()
            AddRs("SelectedDate") = Now
            AddRs("SelectedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        End If
        
        If LocalTimeStr >= "14:00:00" And LocalTimeStr <= "21:59:59" Then
            AddRs("SelectedTimeSlot") = 2
        ElseIf LocalTimeStr >= "06:00:00" And LocalTimeStr <= "13:59:59" Then
            AddRs("SelectedTimeSlot") = 1
        Else
            AddRs("SelectedTimeSlot") = 3
        End If
       
        AddRs("UserId") = Cint(Session("UserId"))
        AddRs("CreatedByName") = Trim(Session("UserName"))
        
        If Request.Form("chkResolved") = "on" Then  
            AddRs("Resolved") = 1
            AddRs("ResolvedBy") = Cint(Session("UserId"))
            AddRs("ResolvedDate") = Date
            AddRs("ResolvedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        Else
            AddRs("Resolved") = 0
        End If
        
        AddRs.Update
        AddRs.MoveLast
        
        NewRecordId = AddRs("Id")
        
        '## Check folder created & has something in it
        If ImgChk(Session("AddFolder")) = True Then
            Dim Fso
            Set Fso = CreateObject("Scripting.FileSystemObject")
            Fso.CopyFolder Session("HsImagePath") & Session("AddFolder"),Session("HsImagePath") & NewRecordId
            Fso.DeleteFolder Session("HsImagePath") & Session("AddFolder"),True
            UpdateImageStatus 1, NewRecordId
        Else
            UpdateImageStatus 0, NewRecordId
        End If   
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing
        
        If Err = 0 Then            
            Session("OrderOk") = "New record " & NewRecordId & " added"
            UpdateReasons
            RedirectUrl = "HsEmail.asp"
            Session("EmailId") = NewRecordId
        Else
            Session("SystemError") = "HS Add Log page"
            RedirectUrl = "SystemError.asp"     
        End If             
    
'############################## Edit Log Update ############################## 
    
    ElseIf Request.Form("frmName") = "frmEditLog" Then
    
        EditRecordId = Request.Form("hLogId")
        UpdateSql = "SELECT * FROM Logs Where (Id = " & EditRecordId & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnHSReports")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        
        UpdateRs("Severity") = Cint(Request.Form("hSeverity"))
        
       ' If Cint(Request.Form("hLocation")) > 0 Then UpdateRs("LocationId") = Cint(Request.Form("hLocation"))
        
        UpdateRs("LastModifiedBy") = Cint(Session("UserId"))       
        UpdateRs("LastModifiedDate") = Date
        UpdateRs("LastModifiedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        
        If Request.Form("hRiddor") > 0 Then
            UpdateRs("Riddor") = Cbool(Request.Form("hRiddor"))
            UpdateRs("RiddorLevel") = Cint(Request.Form("hRiddor"))
            UpdateRs("RiddorDays") = Cint(Request.Form("hRiddorDays"))
        End If
        
        If Request.Form("hRiddor") > 0 Then
            If Request.Form("hRiddorDateSerial") > 0 Then
                '## Already done don't overwrite
            Else 
                UpdateRs("RiddorDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
            End If
        Else
            UpdateRs("RiddorDateSerial") = 0
        End If
        
        If Cbool(Request.Form("hHasNotes")) = False Then        
            If Request.Form("txtResponse") <> "" Then        
                '## need to do formatting so vbcrlf = <br />            
                Notes = Trim(Cstr(Request.Form("txtResponse")))
                Notes = Replace(Notes,VbCrLf,"<br />",1,-1,1)
                '## If triple line feeds cut one out
                '## They seem to like having lots of gaps, making the notes to long
                '## Triple = 1 for return of 1st line then 2 gaps before next line
                Notes = Replace(Notes,"<br /><br /><br />","<br /><br />",1,-1,1)
                Notes = Replace(Notes,"&","&amp;",1,-1,1)
                Notes = Notes & "<br />" & Session("UserName") & " " & Date
                SaveNotes = Trim(Notes) & "<br />"
            End If
        Else
            If Request.Form("txtResponseNew") <> "" Then        
                '## need to do formatting so vbcrlf = <br />            
                NewNotes = Trim(Cstr(Request.Form("txtResponseNew")))
                NewNotes = Replace(NewNotes,VbCrLf,"<br />",1,-1,1)
                '## If triple line feeds cut one out
                '## They seem to like having lots of gaps, making the notes to long
                '## Triple = 1 for return of 1st line then 2 gaps before next line
                NewNotes = Replace(NewNotes,"<br /><br /><br />","<br /><br />",1,-1,1)
                NewNotes = Replace(NewNotes,"&","&amp;",1,-1,1)
                NewNotes = NewNotes & "<br />" & Session("UserName") & " " & Date 
                SaveNotes = "<br />" & Trim(NewNotes) & "<br />"          
            End If
        End If
        
        If Cbool(Request.Form("hHasNotes")) = False Then 
            If SaveNotes <> "" Then        
                UpdateRs("ResponseNotes") = SaveNotes
            End If
        Else
            If SaveNotes <> "" Then 
                UpdateRs("ResponseNotes") = UpdateRs("ResponseNotes") & SaveNotes
            End If
        End If
        
        If UpdateRs("ResolvedDateSerial") = 0 And Request.Form("chkResolved") = "on" Then
            UpdateRs("ResolvedBy") = Cint(Session("UserId"))
            UpdateRs("ResolvedDate") = Date
            UpdateRs("ResolvedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))            
        End If
                                       
        If Request.Form("chkResolved") = "on" Then  
            UpdateRs("Resolved") = 1
        Else
            UpdateRs("Resolved") = 0
        End If  
        
        If ImgChk(Session("EditFolder")) = True Then
            UpdateRs("HasImage") = 1
        Else
            UpdateRs("HasImage") = 0
        End If      
               
        UpdateRs.Update
        UpdateRs.MoveLast 
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing   
        
        If Err = 0 Then
            Session("OrderOk") = "Record " & EditRecordId & " Updated"
        Else
            Session("SystemError") = "HS Edit Log page"
            RedirectUrl = "SystemError.asp"
        End If       
        
         '################################### Clear record lock         
         '## Done when return to option Page Or Login
        
    Else
        Response.Redirect (RedirectUrl)
    End If
End If

Response.Redirect (RedirectUrl)

'#########################################################################

Sub UpdateReasons ()

    Dim UpdateReasonsRs
       
    Set UpdateReasonsRs = Server.CreateObject("ADODB.Recordset")
    UpdateReasonsRs.ActiveConnection = Session("ConnHSReports") 
    UpdateReasonsRs.Source = "Select Id, HasLogs From Reasons Where (Id = " & UpdateReasonId & ")"
    
    UpdateReasonsRs.CursorType = Application("adOpenStatic")
    UpdateReasonsRs.CursorLocation = Application("adUseClient")
    UpdateReasonsRs.LockType = Application("adLockOptimistic")
    UpdateReasonsRs.Open

    UpdateReasonsRs("HasLogs") = 1
    UpdateReasonsRs.Update
    
    UpdateReasonsRs.Close
    Set UpdateReasonsRs = Nothing    

End Sub

'######################################################################

Sub UpdateImageStatus(Status,RecId)

    '## Updates log on new records

    Dim UpdateImageRs
       
    Set UpdateImageRs = Server.CreateObject("ADODB.Recordset")
    UpdateImageRs.ActiveConnection = Session("ConnHSReports") 
    UpdateImageRs.Source = "Select Id, HasImage From Logs Where (Id = " & RecId & ")"
    
    UpdateImageRs.CursorType = Application("adOpenStatic")
    UpdateImageRs.CursorLocation = Application("adUseClient")
    UpdateImageRs.LockType = Application("adLockOptimistic")
    UpdateImageRs.Open

    UpdateImageRs("HasImage") = Status
    
    UpdateImageRs.Update
    
    UpdateImageRs.Close
    Set UpdateImageRs = Nothing

End Sub

'#######################################################################

Function ImgChk(Id)
    
    '## Check if Record has images
    Dim Path
    Path = "C:\Web Sites\MC Reporting\HsImages\" & Cstr(Id)
    Dim ReturnValue
    ReturnValue = False    

    Dim objFSO 
    'Dim objFile
    'Dim Folder
    Dim objFileItem
    Dim objFolder
    Dim objFolderContents
    Dim TotalPics
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(Path) = False Then
        ReturnValue = False
        Set objFSO = Nothing
    Else
        Set objFolder = objFSO.GetFolder(Path)
        Set objFolderContents = objFolder.Files      
        
        TotalPics = 0
        For Each objFileItem in objFolderContents
            If Ucase(Right(objFileItem.Name,4))=".GIF" OR Ucase(Right(objFileItem.Name,4))=".JPG" OR Ucase(Right(objFileItem.Name,4))=".PNG" Then
                TotalPics = TotalPics + 1
            End if
        Next
        
        If TotalPics > 0 Then            
            ReturnValue = True
        Else
            ReturnValue = False
            objFSO.DeleteFolder(Path)
        End If
        
        Set objFSO = Nothing
        Set objFolder = Nothing
        Set objFolderContents = Nothing    
    End If

    ImgChk = ReturnValue

End Function
%>
