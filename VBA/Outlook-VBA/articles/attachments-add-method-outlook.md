---
title: Attachments.Add Method (Outlook)
keywords: vbaol11.chm176
f1_keywords:
- vbaol11.chm176
ms.prod: outlook
api_name:
- Outlook.Attachments.Add
ms.assetid: e11980fd-e1fc-a0c3-cdd0-0e598988d3c2
ms.date: 06/08/2017
---


# Attachments.Add Method (Outlook)

Creates a new attachment in the  **[Attachments](attachments-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Source_** , **_Type_** , **_Position_** , **_DisplayName_** )

 _expression_ A variable that represents an **Attachments** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The source of the attachment. This can be a file (represented by the full file system path with a file name) or an Outlook item that constitutes the attachment.|
| _Type_|Optional| **Long**|The type of the attachment. Can be one of the  **[OlAttachmentType](olattachmenttype-enumeration-outlook.md)** constants.|
| _Position_|Optional| **Long**|This parameter applies only to e-mail messages using the Rich Text format: it is the position where the attachment should be placed within the body text of the message. A value of 1 for the  _Position_ parameter specifies that the attachment should be positioned at the beginning of the message body. A value 'n' greater than the number of characters in the body of the e-mail item specifies that the attachment should be placed at the end. A value of 0 makes the attachment hidden.|
| _DisplayName_|Optional| **String**|This parameter applies only if the mail item is in Rich Text format and  _Type_ is set to **olByValue** : the name is displayed in an **Inspector** object for the attachment or when viewing the properties of the attachment. If the mail item is in Plain Text or HTML format, then the attachment is displayed using the file name in the _Source_ parameter.|

### Return Value

An  **[Attachment](attachment-object-outlook.md)** object that represents the new attachment.


## Remarks

When an  **Attachment** is added to the **Attachments** collection of an item, the **Type** property of the **Attachment** will always return **olOLE** (6) until the item is saved. To ensure consistent results, always save an item before adding or removing objects in the **Attachments** collection.


## Example

The following Microsoft Visual Basic /Visual Basic for Applications (VBA) example creates a mail item, adds an attachment by embedding it at the beginning of the message body, and displays it. To run this example, make sure the attachment which is a file called Test.Doc exists in the C:\ folder.


```vb
Sub AddAttachment() 
 Dim myItem As Outlook.MailItem 
 Dim myAttachments As Outlook.Attachments 
 
 Set myItem = Application.CreateItem(olMailItem) 
 Set myAttachments = myItem.Attachments 
 myAttachments.Add "C:\Test.doc", _ 
 olByValue, 1, "Test" 
 myItem.Display 
End Sub
```


## See also


#### Concepts


[Attachments Object](attachments-object-outlook.md)
#### Other resources


[Attach a File to a Mail Item](http://msdn.microsoft.com/library/1d94629b-e713-92cb-32de-c8910612e861%28Office.15%29.aspx)
[Attach an Outlook Contact Item to an Email Message](http://msdn.microsoft.com/library/ae5240ad-dc3e-4499-8fd0-d8c2d90aa9ba%28Office.15%29.aspx)
[Limit the Size of an Attachment to an Outlook Email Message](http://msdn.microsoft.com/library/9a240e17-f715-482c-9a8b-c6be1144e15a%28Office.15%29.aspx)
[Modify an Attachment of an Outlook Email Message](http://msdn.microsoft.com/library/f5dac09a-272b-49d6-bf1e-82c3981260ed%28Office.15%29.aspx)


