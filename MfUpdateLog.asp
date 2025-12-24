<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp?uf=mf")

Dim Error
Dim UserId
Dim UpdateOk
Dim Count
Dim Notes
Dim UpdateMachineId

Dim AddRs
Dim AddSql
Dim UpdateRs
Dim UpdateSql
Dim RedirectUrl
Dim CreatedDateSerial
Dim EditRecordId
Dim NewRecordId
Dim FaultChk
FaultChk = ""


RedirectUrl = "MfSelectOption.asp"

If Request.Form("chkUpdate") = "on" Then
    '## Testing don't save data, will only be on if at home or on my pc at work
Else

    UpdateMachineId = Cint(Request.Form("hMachine"))
    
'############################## Add New Log ##############################
    If Request.Form("frmName") = "frmAddLog" Then  

    '############################## Create log 

        AddSql = "SELECT * FROM Logs" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnMachinefaults")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open

        AddRs.AddNew
                
        AddRs("MachineId") = Cint(Request.Form("hMachine"))
        AddRs("MachineTypeId") = Cint(Request.Form("hType"))
        AddRs("Status") = Cint(Request.Form("hSeverity"))
        'AddRs("ErrorDescription") = Cstr(Trim(Request.Form("txtDesc")))
        
       
        
        FaultChk = Cstr(Trim(Request.Form("hGroupId")))
        If Left(FaultChk, 1) = "," Then FaultChk = Mid(FaultChk, 2)
        
        AddRs("FaultGroups") = Cstr(Trim(FaultChk))
        AddRs("SelectId") = Cstr(Trim(Request.Form("hSelectId")))
        
        If Request.Form("txtError") <> "" Then        
            '## need to do formatting so vbcrlf = <br />            
            Notes = Trim(Cstr(Request.Form("txtError")))
            Notes = Replace(Notes,VbCrLf,"<br />",1,-1,1)
            '## If triple line feeds cut one out
            '## They seem to like having lots of gaps, making the notes to long
            '## Triple = 1 for return of 1st line then 2 gaps before next line
            Notes = Replace(Notes,"<br /><br /><br />","<br /><br />",1,-1,1)
            Notes = Replace(Notes,"&","&amp;",1,-1,1)
            AddRs("ErrorNotes") = Trim(Notes)
        End If
        
        AddRs("LogDate") = Date
        AddRs("LogDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
       
        AddRs("UserId") = Cint(Session("UserId"))        
        AddRs("CreatedByName") = Trim(Session("UserName"))
        
        If Request.Form("chkRecurring") = "on" Then  
            AddRs("RecurringFault") = 1
        Else
            AddRs("RecurringFault") = 0
        End If
        
        If Request.Form("chkFixed") = "on" Then  
            AddRs("FaultRepaired") = 1
        Else
            AddRs("FaultRepaired") = 0
        End If
        
        AddRs("Cost") = 0
        
        AddRs.Update
        AddRs.MoveLast
        
        NewRecordId = AddRs("Id")
        
        '## Check folder created & has something in it
        If ImgChk(Session("AddFolder")) = True Then
            Dim Fso
            Set Fso = CreateObject("Scripting.FileSystemObject")
            Fso.CopyFolder Session("MfImagePath") & Session("AddFolder"),Session("MfImagePath") & NewRecordId
            Fso.DeleteFolder Session("MfImagePath") & Session("AddFolder"),True
            UpdateImageStatus 1, NewRecordId
        Else
            UpdateImageStatus 0, NewRecordId
        End If        
        
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing
               
        If Err = 0 Then
            UpdateOk = True
            UpdateMachine UpdateMachineId
            Session("OrderOk") = "New record " & NewRecordId & " added"
            If Cint(Request.Form("hSeverity")) = 3 Or Cint(Request.Form("hSeverity")) = 2 Then
                RedirectUrl = "MfEmail.asp"
                Session("EmailId") = NewRecordId                                
            End If
        Else
            Session("SystemError") = "Add MF Log page"
            Response.Redirect ("SystemError.asp")       
        End If
           
'############################## Edit Log Update ############################## 
    
    ElseIf Request.Form("frmName") = "frmEditLog" Then
    
        EditRecordId = Request.Form("hLogId")
        UpdateSql = "SELECT * FROM Logs Where (Id = " & EditRecordId & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnMachinefaults")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        UpdateRs("Status") = Cint(Request.Form("hSeverity"))
        UpdateRs("RepairUserId") = Cint(Session("UserId"))
                
        UpdateRs("RepairDate") = Date
        UpdateRs("RepairDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        
        If Request.Form("txtRepair") <> "" Then        
            '## need to do formatting so vbcrlf = <br />            
            Notes = Trim(Cstr(Request.Form("txtRepair")))
            Notes = Replace(Notes,VbCrLf,"<br />",1,-1,1)
            '## If triple line feeds cut one out
            '## They seem to like having lots of gaps, making the notes to long
            '## Triple = 1 for return of 1st line then 2 gaps before next line
            Notes = Replace(Notes,"<br /><br /><br />","<br /><br />",1,-1,1)
            Notes = Replace(Notes,"&","&amp;",1,-1,1)
            UpdateRs("RepairNotes") = Trim(Notes)
        End If
        
        If Request.Form("chkRecurring") = "on" Then  
            UpdateRs("RecurringFault") = 1
        Else
            UpdateRs("RecurringFault") = 0
        End If
        
        If Request.Form("chkFixed") = "on" Then  
            UpdateRs("FaultRepaired") = 1
        Else
            UpdateRs("FaultRepaired") = 0
        End If
        
        '## Only save fault groups if new selection added
        If Request.Form("hGroupId") <> "" Then
            FaultChk = Cstr(Trim(Request.Form("hGroupId")))
            If Left(FaultChk, 1) = "," Then FaultChk = Mid(FaultChk, 2)
            UpdateRs("FaultGroups") = Cstr(Trim(FaultChk))
            UpdateRs("SelectId") = Cstr(Trim(Request.Form("hSelectId")))
        End If
        
        If CCur(Request.Form("txtCost")) >= 0 Then
            UpdateRs("Cost") = CCur(Request.Form("txtCost"))
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
            UpdateOk = True
            Session("OrderOk") = "Record " & EditRecordId & " Updated"
        Else
            Session("SystemError") = "Edit MF Log page"
            Response.Redirect ("SystemError.asp")       
        End If        
        
         '################################### Clear record lock         
         '## Done when return to option Page Or Login       
        
    Else
        Response.Redirect (RedirectUrl)
    End If
End If

Response.Redirect (RedirectUrl)

'#########################################################################

Sub UpdateMachine (Id)

    Dim UpdateMachineRs
       
    Set UpdateMachineRs = Server.CreateObject("ADODB.Recordset")
    UpdateMachineRs.ActiveConnection = Session("ConnMachinefaults") 
    UpdateMachineRs.Source = "Select Id, HasLogs From Machine Where (Id = " & Id & ")"
    
    UpdateMachineRs.CursorType = Application("adOpenStatic")
    UpdateMachineRs.CursorLocation = Application("adUseClient")
    UpdateMachineRs.LockType = Application("adLockOptimistic")
    UpdateMachineRs.Open

    UpdateMachineRs("HasLogs") = 1
    UpdateMachineRs.Update
    
    UpdateMachineRs.Close
    Set UpdateMachineRs = Nothing    

End Sub

'######################################################################

Sub UpdateImageStatus(Status,RecId)

    '## Updates log on new records

    Dim UpdateImageRs
       
    Set UpdateImageRs = Server.CreateObject("ADODB.Recordset")
    UpdateImageRs.ActiveConnection = Session("ConnMachinefaults") 
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
    Path = "C:\Web Sites\MC Reporting\FaultImages\" & Cstr(Id)
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
