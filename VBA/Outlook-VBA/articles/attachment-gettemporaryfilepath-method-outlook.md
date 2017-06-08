---
title: Attachment.GetTemporaryFilePath Method (Outlook)
keywords: vbaol11.chm3522
f1_keywords:
- vbaol11.chm3522
ms.prod: outlook
api_name:
- Outlook.Attachment.GetTemporaryFilePath
ms.assetid: 3313582b-6241-7a59-0c03-b8af36a17d3d
ms.date: 06/08/2017
---


# Attachment.GetTemporaryFilePath Method (Outlook)

Returns the full path to the attached file that is in a temporary files folder. Read-only.


## Syntax

 _expression_ . **GetTemporaryFilePath**

 _expression_ A variable that represents an **[Attachment](attachment-object-outlook.md)** object.


### Return Value

Returns a  **String** that represents the full path to the temporary attachment file.


## Remarks

The  **GetTemporaryFilePath** method is only valid for those attachments whose **[Type](attachment-type-property-outlook.md)** property is **OlAttachmentType.olByValue** . That means that the attachment is a copy and that the copy can be accessed even if the original file is removed. For other attachment types, the **GetTemporaryFilePath** method returns an error.

 **GetTemporaryFilePath** also returns an error when accessing an **[Attachment](attachment-object-outlook.md)** object in an **[Attachments](attachments-object-outlook.md)** collection or in the **[AttachmentSelection](attachmentselection-object-outlook.md)** object. Use **GetTemporaryFilePath** only in attachment event callbacks listed below for various Microsoft Outlook items:


-  **AttachmentAdd**
    
-  **AttachmentRead**
    
-  **AttachmentRemove**
    
-  **BeforeAttachmentAdd**
    
-  **BeforeAttachmentPreview**
    
-  **BeforeAttachmentRead**
    
-  **BeforeAttachmentSave**
    
-  **BeforeAttachmentWriteToTempFile**
    



## See also


#### Concepts


[Attachment Object](attachment-object-outlook.md)

