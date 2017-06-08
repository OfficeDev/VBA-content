---
title: Attachment.PictureSizeMode Property (Access)
keywords: vbaac10.chm13915
f1_keywords:
- vbaac10.chm13915
ms.prod: access
api_name:
- Access.Attachment.PictureSizeMode
ms.assetid: 07d268ad-d4ba-c9ba-1ef4-7b3e7911ebba
ms.date: 06/08/2017
---


# Attachment.PictureSizeMode Property (Access)

You can use the  **PictureSizeMode** property to specify how a picture for an attachment control is sized. Read/write **Byte**.


## Syntax

 _expression_. **PictureSizeMode**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **PictureSizeMode** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Clip|0|(Default) The picture is displayed in its actual size. If the picture is larger than the attachment control, then the picture is clipped.|
|Stretch|1|The picture is stretched horizontally and vertically to fill the entire attachment control, even if its original ratio of height to width is distorted.|
|Zoom|3|The picture is enlarged to the maximum extent possible while keeping its original ratio of height to width.|
When a small picture is used for the  **DefaultPicture** property of an attachment control, setting the **PictureSizeMode** property to Stretch or Zoom can cause substantial distortion of its resolution. Smaller pictures can be tiled across the entire attachment control by using the **PictureTiling** property.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

