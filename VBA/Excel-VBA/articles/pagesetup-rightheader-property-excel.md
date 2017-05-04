---
title: PageSetup.RightHeader Property (Excel)
keywords: vbaxl10.chm473100
f1_keywords:
- vbaxl10.chm473100
ms.prod: EXCEL
api_name:
- Excel.PageSetup.RightHeader
ms.assetid: 97e1780d-d511-d433-0e31-501381e6318d
---


# PageSetup.RightHeader Property (Excel)

Returns or sets the right part of the header. Read/write  **String** .


## Syntax

 _expression_ . **RightHeader**

 _expression_ A variable that represents a **PageSetup** object.


## Example

This example prints the file name in the upper-right corner of every page.


```vb
Worksheets("Sheet1").PageSetup.RightHeader = "&;F"
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

