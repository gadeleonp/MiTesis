<%
'most control methods throw an exception if an error occurs so we will use an
' On Error Resume Next statement for error trapping purposes
On Error Resume Next
'create an instance of the Upload control
Set objUpload = Server.CreateObject ("Dundas.Upload.2")
'we will save all uploaded data (form data as well as any uploaded files) to disk.
'note that we could also save the uploaded files to memory with the SaveToMemory method.
'also note that both the Save and SaveToMemory methods populate the Form and
' Files collections with one method call.
'its also important to realize that by default files saved to disk will have unique filenames, but
' this can be changed with the UseUniqueNames property

nombrefolder = "Transicion"

RootFolderName = Server.MapPath("..") & "\"
TempFolderName = RootFolderName & "Noticias\" & nombrefolder


objupload.UseUniqueNames = false

objUpload.Save TempFolderName

'check to see if method call was successful using VBScript's Err object, if 
' an error occurred we will redirect user to a fictitious error page
If Err.Number <> 0 Then
Response.Redirect "Error.asp"
Else
folder = objupload.Form("folder")
FinalFolderName = RootFolderName & "Noticias\" & folder

For Each objFormItem In objUpload.Form
Response.Write "<br>The name of the form item is: " & objFormItem
Response.Write "<br>The value of the form item is: " & objFormItem.Value & "<br>"
Next
End If
'use a For Each loop and check to see if the uploaded file is an
' executable (utilizing VBScript's InStr method), if it is delete it from disk.
'but first we will output the name of the file input box(es) responsible for uploads
For Each objUploadedFile in objUpload.Files
Response.Write "The &quot;" & objUploadedFile.TagName & "&quot; file input box was used to upload a file.<br>"
Response.write "Mirame = " & objUploadedFile.value
If InStr(1,objUploadedFile.ContentType,"octet-stream") Then
objUploadedFile.Delete
End If
Next

For Each objuploadedfile in objupload.Files 

objUploadedFile.Move FinalFolderName ,false
Next 
'we will just output the name of the populated form elements and their values

'Release resources
Set objUpload = Nothing
%>
