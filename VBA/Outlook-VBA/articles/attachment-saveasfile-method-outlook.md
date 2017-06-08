---
title: Attachment.SaveAsFile Method (Outlook)
keywords: vbaol11.chm2373
f1_keywords:
- vbaol11.chm2373
ms.prod: outlook
api_name:
- Outlook.Attachment.SaveAsFile
ms.assetid: d755261a-d551-f3a1-1b20-d87d4d9c38ae
ms.date: 06/08/2017
---


# Attachment.SaveAsFile Method (Outlook)

Saves the attachment to the specified path.


## Syntax

 _expression_ . **SaveAsFile**( **_Path_** )

 _expression_ A variable that represents an **Attachment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The location at which to save the attachment.|

## Example

This Visual Basic for Applications (VBA) example uses the  **[SaveAsFile](attachment-saveasfile-method-outlook.md)** method to save the first attachment of the currently open item as a file in the Documents folder, using the attachment's display name as the file name.


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
 
 strPrompt = "Are you sure you want to save the first attachment in the current item to the Documents folder? If a file with the same name already exists in the destination folder, it will be overwritten with this copy of the file." 
 
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

