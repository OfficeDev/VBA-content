---
title: Attachment.DefaultPictureType Property (Access)
keywords: vbaac10.chm10452
f1_keywords:
- vbaac10.chm10452
ms.prod: access
api_name:
- Access.Attachment.DefaultPictureType
ms.assetid: 77032908-5b98-7072-1e53-520485580746
ms.date: 06/08/2017
---


# Attachment.DefaultPictureType Property (Access)

Gets or sets the method used to store the image specified by the  **[DefaultPicture](attachment-defaultpicture-property-access.md)** property in the database. Read/write **Byte**.


## Syntax

 _expression_. **DefaultPictureType**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **DefaultPictureType** property uses the following settings.



|**Setting**|**Value**|**Meaning**|
|:-----|:-----|:-----|
|Embedded (Default)|0|The image is embedded with the specified  **Attachment** control.|
|Linked|1|The image is stored outside of the database.|
|Shared|2|The image is added to the  **[SharedResources](sharedresources-object-access.md)** collection.|

## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

