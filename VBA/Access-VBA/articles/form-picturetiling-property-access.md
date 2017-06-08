---
title: Form.PictureTiling Property (Access)
keywords: vbaac10.chm13383
f1_keywords:
- vbaac10.chm13383
ms.prod: access
api_name:
- Access.Form.PictureTiling
ms.assetid: 9343925c-8184-e9fc-ed62-a272a0bfa0a6
ms.date: 06/08/2017
---


# Form.PictureTiling Property (Access)

You can use the  **PictureTiling** property to specify whether a background picture is tiled across the entire image control, Form window, form, or page of a report. Read/write **Boolean**.


## Syntax

 _expression_. **PictureTiling**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **PictureTiling** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The picture is tiled.|
|No|**False**|(Default) The picture isn't tiled.|
You can also set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

You can create interesting effects by placing a picture on a form or report and setting the  **PictureTiling** property to Yes. The alignment of the tiled images is affected by the **PictureAlignment** property setting. For example, if the **PictureTiling** property is set to Top Left, tiling begins at the top left of the image control, Form window, or page of a report.

If the  **PictureAlignment** property is set to Form Center and the **PictureTiling** property is set to Yes, the background picture of a form will be tiled across the form, not across the Form window.


## See also


#### Concepts


[Form Object](form-object-access.md)

