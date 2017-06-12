---
title: LineFormat.DashStyle Property (Excel)
keywords: vbaxl10.chm110006
f1_keywords:
- vbaxl10.chm110006
ms.prod: excel
api_name:
- Excel.LineFormat.DashStyle
ms.assetid: b1a6f135-ca68-5399-9156-3044e99bf3ab
ms.date: 06/08/2017
---


# LineFormat.DashStyle Property (Excel)

Returns or sets the dash style for the specified line. Can be one of the  **[MsoLineDashStyle](http://msdn.microsoft.com/library/aba7f9d7-1689-c4a8-3b1e-e8dfb4a81d44%28Office.15%29.aspx)** contants. Read/write **Long** .


## Syntax

 _expression_ . **DashStyle**

 _expression_ A variable that represents a **LineFormat** object.


## Example

This example adds a blue dashed line to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(10, 10, 250, 250).Line 
    .DashStyle = msoLineDashDotDot 
    .ForeColor.RGB = RGB(50, 0, 128) 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-excel.md)

