---
title: Attachment Object (Publisher)
keywords: vbapb10.chm9240575
f1_keywords:
- vbapb10.chm9240575
ms.prod: publisher
api_name:
- Publisher.Attachment
ms.assetid: d617bdf6-b0ba-be0d-0f72-f729010636c1
ms.date: 06/08/2017
---


# Attachment Object (Publisher)

Represents an attachment to a merged e-mail message.


## Remarks

An **Attachment** object corresponds to one of the attachments in the list of attachments in the **Attachments** box in the **Merge to E-mail** dialog box in the Microsoft Publisher user interface. (On the **File** menu, point to **Send E-mail**, click  **Send E-mail Merge**, and then click  **Options**.)

To remove the attachment from the merged e-mail, use the  **Delete** method of the **Attachment** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Add** method to add an attachment to an e-mail merge message. It adds an **Attachment** object that represents a bitmap image to the **Attachments** collection of the active document.

Before running this macro, place a file named  _image.bmp_ in the root of the C drive on your computer, or change the name and file path of the file in the macro to specify the one you want to attach.

Note that to send an e-mail merge message, you must connect to a data source, create the e-mail merge, and then send the message. For more information, see the  **[EmailMergeEnvelope](http://msdn.microsoft.com/library/555dd80e-bac2-96dd-4256-ad1b8006da0f%28Office.15%29.aspx)** object topic.




```
Public Sub Attachment_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 Dim pubAttachment As Publisher.Attachment 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachemts 
 
 Set pubAttachment = pubAttachments.Add("C:\image.bmp ") 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/935fa9e7-9d40-b820-e386-1a1960845da1%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Name](http://msdn.microsoft.com/library/7539a5ac-427f-0dfe-dc31-47ef9436fd14%28Office.15%29.aspx)|

## See also


#### Other resources


[Attachment Object Members](http://msdn.microsoft.com/library/594cf3eb-73d8-afa9-b598-ab68066dde8b%28Office.15%29.aspx)
