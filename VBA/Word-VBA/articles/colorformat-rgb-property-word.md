---
title: ColorFormat.RGB Property (Word)
keywords: vbawd10.chm163971072
f1_keywords:
- vbawd10.chm163971072
ms.prod: word
api_name:
- Word.ColorFormat.RGB
ms.assetid: 78158429-359c-bc6e-9e81-a119aace776c
ms.date: 06/08/2017
---


# ColorFormat.RGB Property (Word)

Returns or sets the red-green-blue (RGB) value of the specified color. Read/write  **Long** .


## Syntax

 _expression_ . **RGB**

 _expression_ A variable that represents a **[ColorFormat](colorformat-object-word.md)** object.


## Example

This example sets the color of the second shape in the active document to gray.


```vb
ActiveDocument.Shapes(2).Fill.ForeColor.RGB = RGB(128, 128, 128)
```

This example sets the color of the shadow for Rectangle 1 in the active document to blue.




```vb
ActiveDocument.Shapes("Rectangle 1").Shadow.ForeColor.RGB = _ 
 RGB(0, 0, 255)
```

This example returns the value of the foreground color of the first shape in the active document.




```vb
MsgBox ActiveDocument.Shapes(1).Fill.ForeColor.RGB
```


## See also


#### Concepts


[ColorFormat Object](colorformat-object-word.md)

