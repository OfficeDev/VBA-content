---
title: PageSetup.Orientation Property (Excel)
keywords: vbaxl10.chm473090
f1_keywords:
- vbaxl10.chm473090
ms.prod: EXCEL
api_name:
- Excel.PageSetup.Orientation
ms.assetid: 9e41d5c8-e887-3212-c298-c2921137ec9c
---


# PageSetup.Orientation Property (Excel)

Returns or sets a  **[XlPageOrientation](xlpageorientation-enumeration-excel.md)** value that represents the portrait or landscape printing mode.


## Syntax

 _expression_ . **Orientation**

 _expression_ A variable that represents a **PageSetup** object.


## Example

This example sets Sheet1 to be printed in landscape orientation.


```vb
Worksheets("Sheet1").PageSetup.Orientation = xlLandscape
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

