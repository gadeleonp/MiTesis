<%@ Language=VBScript %>
<% Response.Buffer = true
'******************************************************************************
'Copyright Dundas Software Ltd. 2000. All Rights Reserved.
'
'PURPOSE:			This page processes the data POSTED by main.asp.
'					If the upload operation is successful the user will receive
'					  a success message, if the operation fails then a failure
'					  message will be displayed.  These messages will be passed
'					  back to main.asp via a QueryString.	    
'
'POST DESTINATION:  NA           
'
'COMMENTS:			Note that a maximum size of one (1) MBytes per file is
'					  allowed.
'
'					Uploaded files will be saved to memory.			
'
'Dundas Software Contact Information:
'	Email:	sales@dundas.com
'   Phone:	(800) 463-1492
'			(416) 467-5100
'	Fax:	(416) 422-4801
'******************************************************************************

'most methods will throw an exception if the method is unsuccessful so we will
'  enable inline error trapping.
on error resume next

dim objUpload		'instance of Uplaod control
dim strMessage		'stores success/failure message sent back to main.asp
dim strPath		 'set to "c:\4AA80F25-21E4-11D4-9985-0050BAD44BCD", it is the path of the storage folder

'create an instance of the Upload Control and trap for object creation failure
set objUpload = server.CreateObject("Dundas.Upload.2")
if  err.number <> 0 then
	Response.Redirect "main.asp?Message=" & err.description
end if


'set maximum file size allowed to approx. 1 MBytes
objUpload.MaxFileSize = 1048576

'save all uploaded form data to memory.  Note that this also populates the Files and
'	Form collections with ALL uploaded form data
strPath = Server.MapPath("..")
strPath = strPath & "\" & "00566F20-168D-445E-974E-A5BC0881F6A4"

objUpload.DirectoryCreate strPath  'create temp directory if it doesn't exist

objUpload.Save strPath 

'objUpload.SaveToMemory

'now trap for success/failure of operation, and also use the control's Form collection
'	to retrieve the name entered by the user so we can send his/her name back to main.asp
dim temp
if IsEmpty(objUpload.Form("txtName")) = false then temp = " "
if err.number <> 0 then
	strMessage = "Sorry " & objUpload.Form("txtName") & temp & "but the following error occurred: " & err.description
else
	strMessage = "The upload operation was successfully performed" & temp & objUpload.Form("txtName") & "."
end if	

'now use a response.redirect to get user back to main.asp
Response.Redirect "main.asp?Message=" & strMessage

'release resources
set objUpload = nothing
%>