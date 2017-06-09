---
title: Image.PictureAlignment Property (Access)
keywords: vbaac10.chm10370
f1_keywords:
- vbaac10.chm10370
ms.prod: access
api_name:
- Access.Image.PictureAlignment
ms.assetid: e0ebec64-9a26-859e-b9fd-5f4a47253bba
ms.date: 06/08/2017
---


# Image.PictureAlignment Property (Access)

You can use the  **PictureAlignment** property to specify where a background picture will appear in an image control or on a form or report. Read/write **Byte**.Read/write.


## Syntax

 _expression_. **PictureAlignment**

 _expression_ A variable that represents an **Image** object.


## Remarks

The  **PictureAlignment** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Top Left|0|The picture is displayed in the top-left corner of the image control, Form window, or page of a report.|
|Top Right|1|The picture is displayed in the top-right corner of the image control, Form window, or page of a report.|
|Center|2|(Default) The picture is centered in the image control, Form window, or page of a report.|
|Bottom Left|3|The picture is displayed in the bottom-left corner of the image control, Form window, or page of a report.|
|Bottom Right|4|The picture is displayed in the bottom-right corner of the image control, Form window, or page of a report.|
|Form Center|5|(Forms only) The form's picture is centered horizontally in relation to the width of the form and vertically in relation to the height the entire form.|
You can also set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

This property can be set in any view.

The Form Center setting aligns a form's picture in the center of the form itself. All other  **PictureAlignment** property settings align a form's picture in relation to the Form window. If you want to make sure that a form's picture is displayed only on the form or tiled across only the form, set the **PictureAlignment** property to Form Center.

For reports, the picture appears relative to a full page and not in relation to the size of the actual report. If your report is less than a full page and you want a picture to appear at a location not available through the  **PictureAlignment** property settings, use an image control instead.

When you set the  **PictureTiling** property to Yes, tiling of the picture will begin from the **PictureAlignment** property setting.


## Example

The following example displays the picture "Logo.gif" in the top left corner of the "Purchase Order" report.


```vb
With Reports("Purchase Order") 
 .Picture = "C:\Picture Files\Logo.gif" 
 .PictureAlignment = 0 
End With
```


## See also


#### Concepts


[Image Object](image-object-access.md)

