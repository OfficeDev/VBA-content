---
title: Shape.PictureFormat Property (Word)
keywords: vbawd10.chm161480822
f1_keywords:
- vbawd10.chm161480822
ms.prod: word
api_name:
- Word.Shape.PictureFormat
ms.assetid: 638513d0-e40b-c220-1c56-72c1160afada
ms.date: 06/08/2017
---


# Shape.PictureFormat Property (Word)

Returns a  **PictureFormat** object that contains picture formatting properties for the specified object. Read-only.


## Syntax

 _expression_ . **PictureFormat**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

This property applies to  **Shape** objects that represent pictures or OLE objects.


## Example

This example sets the brightness and contrast for shape one on  _myDocument_ . Shape one must be a picture or an OLE object.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).PictureFormat 
 .Brightness = 0.3 
 .Contrast = .75 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

