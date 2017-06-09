---
title: OlAttachmentType Enumeration (Outlook)
keywords: vbaol11.chm3052
f1_keywords:
- vbaol11.chm3052
ms.prod: outlook
api_name:
- Outlook.OlAttachmentType
ms.assetid: b6373ef7-0f30-d6c4-eb52-c6ef1de40b52
ms.date: 06/08/2017
---


# OlAttachmentType Enumeration (Outlook)

Specifies the attachment type.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olByReference**|4|This value is no longer supported since Microsoft Outlook 2007. Use  **olByValue** to attach a copy of a file in the file system.|
| **olByValue**|1|The attachment is a copy of the original file and can be accessed even if the original file is removed.|
| **olEmbeddeditem**|5|The attachment is an Outlook message format file (.msg) and is a copy of the original message.|
| **olOLE**|6|The attachment is an OLE document.|

## Remarks

Used as an optional parameter to the [Attachments.Add Method (Outlook)](attachments-add-method-outlook.md) to specify the attachment type.


## See also


#### Other resources


[Attach a File to a Mail Item](http://msdn.microsoft.com/library/1d94629b-e713-92cb-32de-c8910612e861%28Office.15%29.aspx)
[Attach an Outlook Contact Item to an Email Message](http://msdn.microsoft.com/library/ae5240ad-dc3e-4499-8fd0-d8c2d90aa9ba%28Office.15%29.aspx)
[Limit the Size of an Attachment to an Outlook Email Message](http://msdn.microsoft.com/library/9a240e17-f715-482c-9a8b-c6be1144e15a%28Office.15%29.aspx)
[Modify an Attachment of an Outlook Email Message](http://msdn.microsoft.com/library/f5dac09a-272b-49d6-bf1e-82c3981260ed%28Office.15%29.aspx)

