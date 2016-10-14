'define the script parameters
Const strOlFolName = "Internal Audit"
Const strSender = "root"
Const strSubject = "Satellite Inventory Report"
Const strDropLoc = "\\xxxxx"

'script specific parameters
Dim oFolder, cMessages, oMailItem, counter, moveMail(), oAtt, a
Set oOutlook = WScript.CreateObject("Outlook.Application")
Set oNameSpace = oOutlook.GetNameSpace("MAPI")
Set oMailbox = oNameSpace.Folders(strOlFolName)
Set oFolder = oMailbox.Folders("Inbox")
Set cMessages = oFolder.Items
Set moveToFol = oMailbox.Folders("Inbox").Folders("CAP IP Mapping")

counter = 0

For Each oMailItem In cMessages
	If TypeName(oMailItem) = "MailItem" Then
		If oMailItem.Sender = strSender And Left(oMailItem.Subject,26) = strSubject Then
			counter = counter + 1
			ReDim Preserve moveMail(counter)
			Set moveMail(counter) = oMailItem
			Set cAttachments = oMailItem.Attachments
			For Each oAtt In cAttachments
				oAtt.SaveAsFile strDropLoc & "\" & oAtt.FileName
			Next
		End If
	End If
Next

For a = 1 To counter
	moveMail(a).Move moveToFol
Next

