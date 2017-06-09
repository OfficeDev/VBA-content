---
title: Attachment.DisplayName Property (Outlook)
keywords: vbaol11.chm2365
f1_keywords:
- vbaol11.chm2365
ms.prod: outlook
api_name:
- Outlook.Attachment.DisplayName
ms.assetid: 2321da5d-4aae-c483-f41e-03b35af80dd1
ms.date: 06/08/2017
---


# Attachment.DisplayName Property (Outlook)

Returns or sets a  **String** representing the name, which does not need to be the actual file name, displayed below the icon representing the embedded attachment. Read/write.


## Syntax

 _expression_ . **DisplayName**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

This property corresponds to the MAPI property  **PidTagDisplayName** .


## Example

This Visual Basic for Applications (VBA) example uses the  **[Attachment.SaveAsFile](attachment-saveasfile-method-outlook.md)** method to save the first attachment of the currently open item as a file in the Documents folder, using the attachment's display name as the file name.


```vb
Sub SaveAttachment() 
 
 Dim myInspector As Outlook.Inspector 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myAttachments As Outlook.Attachments 
 
 
 
 Set myInspector = Application.ActiveInspector 
 
 If Not TypeName(myInspector) = "Nothing" Then 
 
 If TypeName(myInspector.CurrentItem) = "MailItem" Then 
 
 Set myItem = myInspector.CurrentItem 
 
 Set myAttachments = myItem.Attachments 
 
 'Prompt the user for confirmation 
 
 Dim strPrompt As String 
 
 strPrompt = "Are you sure you want to save the first attachment " &; _ 
 
 "in the current item to the Documents folder? If a file with the " &; _ 
 
 "same name already exists in the destination folder, " &; _ 
 
 "it will be overwritten with this copy of the file." 
 
 If MsgBox(strPrompt, vbYesNo + vbQuestion) = vbYes Then 
 
 myAttachments.Item(1).SaveAsFile Environ("HOMEPATH") &; "\My Documents\" &; _ 
 
 myAttachments.Item(1).DisplayName 
 
 End If 
 
 Else 
 
 MsgBox "The item is of the wrong type." 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Attachment Object](attachment-object-outlook.md)

