---
title: Attachment.PictureAlignment Property (Access)
keywords: vbaac10.chm13916
f1_keywords:
- vbaac10.chm13916
ms.prod: access
api_name:
- Access.Attachment.PictureAlignment
ms.assetid: 505daae0-8321-cce0-028a-ff6c2ac16245
ms.date: 06/08/2017
---


# Attachment.PictureAlignment Property (Access)

You can use the  **PictureAlignment** property to specify where a background picture will appear in the **Attachment** control. Read/write **Byte**.


## Syntax

 _expression_. **PictureAlignment**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **PictureAlignment** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Top Left|0|The picture is displayed in the top-left corner of the control.|
|Top Right|1|The picture is displayed in the top-right corner of the control.|
|Center|2|(Default) The picture is centered in the control.|
|Bottom Left|3|The picture is displayed in the bottom-left corner of the control.|
|Bottom Right|4|The picture is displayed in the bottom-right corner of the control.|
You can also set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

This property can be set in any view.

When you set the  **PictureTiling** property to Yes, tiling of the picture will begin from the **PictureAlignment** property setting.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

