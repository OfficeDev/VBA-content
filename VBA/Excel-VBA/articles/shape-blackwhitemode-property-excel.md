---
title: Shape.BlackWhiteMode Property (Excel)
keywords: vbaxl10.chm636118
f1_keywords:
- vbaxl10.chm636118
ms.prod: excel
api_name:
- Excel.Shape.BlackWhiteMode
ms.assetid: 95a00870-82c2-d193-6971-9f92aeed6532
ms.date: 06/08/2017
---


# Shape.BlackWhiteMode Property (Excel)

Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write  **[MsoBlackWhiteMode](http://msdn.microsoft.com/library/2b4d7e22-1277-9f5c-ba52-a37e113477c1%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **BlackWhiteMode**

 _expression_ A variable that represents a **Shape** object.


## Example

This example sets shape one on  `wksOne` to appear in black-and-white mode. When you view the presentation in black-and-white mode, shape one will appear black regardless of what color it is in color mode.


```vb
Sub UseBlackWhiteMode() 
 
    Dim wksOne As Worksheet 
    Set wksOne = Application.Worksheets(1) 
    wksOne.Shapes(1).BlackWhiteMode = msoBlackWhiteGrayOutline 
 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

