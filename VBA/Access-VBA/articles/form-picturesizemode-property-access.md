---
title: Form.PictureSizeMode Property (Access)
keywords: vbaac10.chm13381
f1_keywords:
- vbaac10.chm13381
ms.prod: access
api_name:
- Access.Form.PictureSizeMode
ms.assetid: b2e7646c-a040-0205-b840-0ed5b43982ab
ms.date: 06/08/2017
---


# Form.PictureSizeMode Property (Access)

You can use the  **PictureSizeMode** property to specify how a picture for a form or report is sized. Read/write **Byte**.


## Syntax

 _expression_. **PictureSizeMode**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **PictureSizeMode** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Clip|0|(Default) The picture is displayed in its actual size. If the picture is larger than the form or report, then the picture is clipped.|
|Stretch|1|The picture is stretched horizontally and vertically to fill the entire form, even if its original ratio of height to width is distorted.|
|Zoom|3|The picture is enlarged to the maximum extent possible while keeping its original ratio of height to width.|
|Stretch Horizontal|4|The picture is stretched horizontally to fit the width of the form.|
|Stretch Vertical|5|The picture is stretched vertically to fit the height of the form.|
When a small picture is used for the  **Picture** property of a form or report, setting the **PictureSizeMode** property to Stretch or Zoom can cause substantial distortion of its resolution. Smaller pictures can be tiled across the entire form or report by using the **PictureTiling** property.


## Example

The following example sets the background picture of the "Order Entry" form to "Contacts.gif", and stretches the picture to fit the entire form's background.


```vb
With Forms("Order Entry") 
 .Picture = "C:\Picture Files\Contacts.gif" 
 .PictureSizeMode = 1 
End With
```


## See also


#### Concepts


[Form Object](form-object-access.md)

