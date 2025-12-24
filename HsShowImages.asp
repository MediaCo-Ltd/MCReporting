<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%

'On Error Resume Next

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

If Session("UserName") = "" Then Response.Redirect ("CloseWindow.asp")

Dim LogToGet
LogToGet = Clng(Request.QueryString("id"))

Dim FolderName
If LogToGet < 10 Then
    FolderName = "0" & Cstr(LogToGet)
Else
    FolderName = Cstr(LogToGet)
End If    

'Dim strCurFile

Dim strImageHtml
'Dim strDisplayText
Dim strPhysicalPath
'Dim strImageCaption
'Dim strImageFooter
Dim strImagePath
Dim strPicArray()
'Dim strFolderArray()

Dim intImageNumber
Dim intImageCount
'Dim intCount
Dim intSubCount
'Dim intFoldercount

Dim objFSO 
Dim objFile
Dim Folder
Dim objFileItem
Dim objFolder
Dim objFolderContents

'Dim TimeExtra
'dim ImageTitle


'## Background = #333333

Dim intTotPics, intPicsPerRow, intPicsPerPage, intTotPages, intPage
'## Set up rows and number of pictures
intPicsPerRow  = 4
intPicsPerPage = 12 '## not realy needed

'strCurFile = "MfShowImages.asp"

'strPhysicalPath = Session("DriveName") & "\Web Sites\MC Reporting\HsImages\"  & FolderName

strPhysicalPath = Session("HsImagePath") & FolderName
strImagePath = "HsImages/" & FolderName & "/"
'On Error Resume Next


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
    <head>
        <title>HS Reporting</title>
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
        <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
        <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
        <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
        <script type="text/javascript" src="JsFiles/HSReportsJSFunc.js"></script>
    </head>
    
    <body style="padding: 0px; margin: 0px" onload="javascript:ShowImageLoadChk();">
        <table style="width: 100%; padding-right: 10px; padding-left: 10px;" >
            <tr>
		        <td align="left" valign="bottom" height="100" colspan="3">
                    <img align="left" alt="mediaco logo" src='Images/mediaco_logo.jpg' width="160"   />
                </td>            
            </tr>
	        <tr>
		        <td height="8" valign="top" colspan="3">
		            <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		        </td>
	        </tr>	
            <tr>
                <td align="left" style="height: 20px" valign="bottom" width="33%" >
                    <a href ="javascript:window.close();" style="font-size:12px; color:<%=NewBlue%>">Close&nbsp;Window</a>&nbsp;&nbsp;
                </td>
                <td height="20px" width="34%" valign="bottom" align="center"> <!-- color #0069AA; Blue-->
                    <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" /> 
                    <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">Images for record <%=LogToGet%></font>
                </td>
                <td valign="bottom" align="right" width="33%">
                    <a href ="javascript:window.close();" style="font-size:12px; color:<%=NewBlue%>">Close&nbsp;Window</a>&nbsp;&nbsp;
                </td> 
            </tr>              
        </table>
        
       
        <br />
      
        
        <center>                   
            <table border="0" cellpadding="0" cellspacing="5" style=" width:96%; border-collapse: collapse;">
                <tr>
                    <%     
                       
                    Set objFSO = CreateObject("Scripting.FileSystemObject")
                    Set objFolder = objFSO.GetFolder(strPhysicalPath)
                    Set objFolderContents = objFolder.Files                      
                    
                    '## Get the number of pictures in the directory 
                    intTotPics = 0
                    For Each objFileItem in objFolderContents
                        If Ucase(Right(objFileItem.Name,4))=".GIF" OR Ucase(Right(objFileItem.Name,4))=".JPG" OR Ucase(Right(objFileItem.Name,5))=".JPEG" OR Ucase(Right(objFileItem.Name,4))=".PNG" Then
                            intTotPics = intTotPics + 1
                        End if
                    Next
                    
                    If intTotPics > 1 Then                             
                        Redim strPicArray(1,intTotPics)
                        '## Store picture file names in an array
                        intSubCount = 0
                        For Each objFileItem in objFolderContents
                            'GIF pictures
                            If Ucase(Right(objFileItem.Name,4))=".GIF" Then			 
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)			
                                intSubCount = intSubCount + 1
                            'JPG pictures
                            ElseIf Ucase(Right(objFileItem.Name,4))=".JPG" Then				
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)
                                intSubCount = intSubCount + 1
                            'JPEG pictures
                            ElseIf Ucase(Right(objFileItem.Name,4))=".JPEG" Then				
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)
                                intSubCount = intSubCount + 1
                            'PNG pictures
                            ElseIf Ucase(Right(objFileItem.Name,4))=".PNG" Then				
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)
                                intSubCount = intSubCount + 1
                            End if
                        Next                            
                            
                        For intImageNumber = 0 To intSubCount -1
                            
                            '## to sort out image placement in different browsers
                            Select Case Session("Browser")
                                Case "Safari"     '## added this to safari <br /> at start'
                                    strImageHtml = "<br /><br /><img style='border: 1px solid #000000' alt='' width = '160' height = '120' src ='" & strImagePath & "/" & strPicArray(0,intImageNumber) & "' /></a><br />"
                                Case "IE", "Opera"  '## Both Ok
                                    strImageHtml = "<br /><br /><img  style='display:block; border: 1px solid #000000' alt='' width = '160' height = '120' src ='" & strImagePath & "/" & strPicArray(0,intImageNumber) & "' /></a><br />"
                                Case "Firefox", "Chrome" '## Ok
                                    strImageHtml = "<br /><br /><img style='border: 1px solid #000000;' alt='' width = '160' height = '120' src ='" & strImagePath & "/" & strPicArray(0,intImageNumber) & "' /></a><br />"
                            End Select  
                                        
                            Response.write "<td align='center'>"
                            Response.write "<a style='color:#333333;' href=""Javascript:ShowSingleImage('" & FolderName & "/" & strPicArray(0,intImageNumber) & "')"">"
                            Response.Write(strImageHtml) 
                        
                            Response.Write "</td>"
                            Err.Clear
                            
                            intImageCount = intImageCount + 1
                            If intImageCount = intPicsPerRow Then
                                Response.write "</tr><tr>"
                                intImageCount = 0
                            End If
                            
                        Next
                    Else
                        '## Single image
                        Redim strPicArray(1,intTotPics)
                        '## Store picture file names in an array
                        intSubCount = 0
                        For Each objFileItem in objFolderContents
                            'GIF pictures
                            If Ucase(Right(objFileItem.Name,4))=".GIF" Then			 
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)			
                                intSubCount = intSubCount + 1
                            'JPG pictures
                            ElseIf Ucase(Right(objFileItem.Name,4))=".JPG" Then				
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)
                                intSubCount = intSubCount + 1
                            'JPEG pictures
                            ElseIf Ucase(Right(objFileItem.Name,5))=".JPEG" Then				
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)
                                intSubCount = intSubCount + 1
                            'PNG pictures
                            ElseIf Ucase(Right(objFileItem.Name,4))=".PNG" Then				
                                strPicArray(0,intSubCount) = objFileItem.Name
                                strPicArray(1,intSubCount) = Cstr(int(intSubCount/intTotPics)+1)
                                intSubCount = intSubCount + 1
                            End if
                        Next
                           
                        '## to sort out image placement in different browsers   
                        Select Case Session("Browser")
                            Case "Safari"     '## added this to safari <br /> at start'
                                strImageHtml = "<br /><br /><img ' alt='' width='750' height='auto' src ='" & strImagePath & "/" & strPicArray(0,intImageNumber) & "' /><br />"
                            Case "IE", "Opera"  '## Both Ok
                                strImageHtml = "<br /><img  style='display:block;' width='750' height='auto' alt=''  src ='" & strImagePath & "/" & strPicArray(0,intImageNumber) & "' />"
                            Case "Firefox", "Chrome" '## Ok
                                strImageHtml = "<br /><br /><img alt='' width='750' height='auto' src ='" & strImagePath & "/" &  strPicArray(0,intImageNumber) & "' /><br/>"
                        End Select                               
                      
                        Response.write "<td align='center'>"
                        Response.Write(strImageHtml) 
                    
                        Response.Write "</td>" 
                    End If
                    
                    Set objFSO = Nothing
                    Set objFile = Nothing
                    Set Folder = Nothing
                    Set objFileItem = Nothing
                    Set objFolder = Nothing
                    Set objFolderContents = Nothing
                    Erase strPicArray
                    
                    Err.Clear                                      
                    %>
                    <td height='50'>&nbsp;</td>                    
                </tr>
            </table>
        </center>
    </body>
</html>






