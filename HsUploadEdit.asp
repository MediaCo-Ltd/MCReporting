<%@Language="VBScript" EnableSessionState=True%>

<!-- #include file="UploadFunctions.asp" -->
<!-- #include file="..\##GlobalFiles\DefaultColours.asp" -->

<%  
'## If want this (option explicit)
'## Includes have to go after it

Response.Expires = -1
Server.ScriptTimeout = 600
' All communication must be in UTF-8, including the response back from the request
Session.CodePage  = 65001

Dim fso



'## Path will allways be as below regardles of what pc its on
Dim UploadDestination
UploadDestination = Session("HsImagePath") & Session("EditFolder")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>Image Upload</title>
    
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
    
    <script type="text/javascript">

        function bodyunload() 
        {

            if (document.getElementById("file").value != '')
             {
                window.opener.location.reload();
            } 

        }     
        
        function onSubmitForm() 
        {
            var filename = document.getElementById("attach").value
            if (filename == "")
            { alert("Please press the Browse button and pick a file."); return false; }
            else
                if (CheckfileName() == true)
                { document.getElementById("loading").innerHTML = "Uploading File, please wait.."
                    return true; }
                else
                    { return false; }
        }

        function CheckfileName() {
            var filename = document.getElementById("attach").value
            if (filename.indexOf(".gif") > 0 || filename.indexOf(".jpg") > 0 || filename.indexOf(".png") > 0 || filename.indexOf(".jpeg") > 0) 
            {
                if (filename.indexOf("&") > 0 || filename.indexOf(" ") > 0 || filename.indexOf("?") > 0 || filename.indexOf("#") > 0 || filename.indexOf("'") > 0) 
                {
                    alert("Rename file & remove invalid characters");
                    return false;
                }
                else
                { return true; }
            }
            else {
                alert("This is not a Image File");
                return false;
            }
        }

        function SetTimeOut() 
        {
            setTimeout("window.close()", 300000);
        }
    </script>
</head>

<body class="NonDataTables" style=" padding: 0px; margin: 0px" onunload="bodyunload();" onload="SetTimeOut();"> 
<form name='frmSend' method='post' enctype='multipart/form-data' accept-charset='utf-8' action='HsUploadEdit.asp' onsubmit='return onSubmitForm();'>

    <table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
        <tr>
	        <td align="left" valign="middle" height="60">
                <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="100" />
            </td>
        </tr>
	    <tr>
		    <td height="8" valign="top" >
		        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		    </td>
	    </tr>
	    <tr>            
            <td align="right" valign="bottom" >
                <a href ="javascript:window.close();" style="color:<%=NewBlue%>">Close&nbsp;window</a>
            </td>                        
        </tr>	        
    </table> 
    
    <table style='width: 100%; position: absolute; top: 30%; padding-right: 10px; padding-left: 30px;' >
        <tr>
            <td >
                If&nbsp;an&nbsp;image&nbsp;already&nbsp;exists.<br />
                It&nbsp;will&nbsp;not&nbsp;be&nbsp;overwritten.<br />
                It&nbsp;will&nbsp;be&nbsp;saved&nbsp;as&nbsp;a&nbsp;duplicate.
                <br /><br />
                <label style="font-size: 12px; color: #FF0000">Rename your file if it has has any of the following..</label> 
                <br />               
                <label style="font-size: 12px; color: #FF0000">&amp; % ? # Apostrophes or spaces in the name</label>
                <br /><br />
            </td>
        </tr>
        <tr>
            <td >Select&nbsp;File</td>
        </tr>        
        <tr>
            <td >
                <input name='attach' id='attach' type='file' size='35' onchange='return CheckfileName();' onmouseup='return false;'/>
            </td>
        </tr>        
        <tr>
            <td align="center">&nbsp;<span style="color:<%=NewBlue%>; font-size: medium;" id="loading" ></span></td>
        </tr>        
        <tr>
            <td >
                <br /><input type='submit' name="btnSubmit" id="btnSubmit" value='Upload' />
            </td>
        </tr>        
    </table>

<%
Dim Diagnostics

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Diagnostics = TestEnvironment()
    If Diagnostics<>"" Then
        Response.Write "<table style='width: 100%; position: absolute; bottom: 15%; padding-right: 10px; padding-left: 30px;' >" & VbCrLf
        Response.Write "    <tr>" & VbCrLf
        Response.Write "        <td>" & VbCrLf
        Response.Write Diagnostics & "<br/>After you correct this problem, reload the page." & VbCrLf
        Response.Write "        </td>" & VbCrLf
        Response.Write "    </tr>" & VbCrLf
        Response.Write "</table>"
    End If
Else
    Response.Write "<table style='width: 100%; position: absolute; bottom: 15%; padding-right: 10px; padding-left: 30px;' >" & VbCrLf
    Response.Write "    <tr>" & VbCrLf
    Response.Write "        <td>" & VbCrLf
    Response.Write "            <span style='color:" & NewBlue & " ; font-size: medium;'>" &  SaveFiles() & ")&nbsp;uploaded</span>" & VbCrLf
    Response.Write "            <input type='hidden' id='file' value='1' />" & VbCrLf
    Response.Write "        </td>" & VbCrLf
    Response.Write "    </tr>" & VbCrLf
    Response.Write "</table>"
End If

Function TestEnvironment()

    Dim fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")   
    
    If Not fso.FolderExists(UploadDestination) Then
        fso.CreateFolder(UploadDestination)        
        Exit Function
    End If
    
    fileName = UploadDestination & "\test.txt"
    On Error Resume Next
    Set testFile = fso.CreateTextFile(fileName, True)
    If Err.Number<>0 Then
        TestEnvironment = "<b>Folder " & UploadDestination & " does not have write permissions."
        Exit Function
    End If
    
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 Then
        TestEnvironment = "<b>Folder " & UploadDestination & " does not have delete permissions."
        Exit Function
    End If
    
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 Then
        TestEnvironment = "<b>The ADODB object <i>Stream</i> is not available in your server."
        Exit Function
    End If
    
    Set streamTest = Nothing
    Set fso = Nothing
    
End Function

Function SaveFiles

    Dim Upload, fileName, fileSize, ks, i, fileKey

    Set Upload = New FreeASPUpload
    Upload.Save(UploadDestination)
    
	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 Then Exit Function

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    If (UBound(ks) <> -1) Then
        SaveFiles = "File&nbsp;("
        For Each fileKey In Upload.UploadedFiles.keys
            SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName 
        Next
    Else
        SaveFiles = "No file selected for upload or the file name specified in the upload form does not correspond to a valid file in the system."
    End If
    
End Function
%>
</form>
</body>
</html>
