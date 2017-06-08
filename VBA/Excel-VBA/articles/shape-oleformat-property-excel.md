---
title: Shape.OLEFormat Property (Excel)
keywords: vbaxl10.chm636130
f1_keywords:
- vbaxl10.chm636130
ms.prod: excel
api_name:
- Excel.Shape.OLEFormat
ms.assetid: 7f2ff868-a7cf-3a9f-4ad8-6213f55573ea
ms.date: 06/08/2017
---


# Shape.OLEFormat Property (Excel)

Returns an  **[OLEFormat](oleformat-object-excel.md)** object that contains OLE object properties. Read-only.


## Syntax

 _expression_ . **OLEFormat**

 _expression_ A variable that represents a **Shape** object.


## Example

This example activates an OLE object. If  `Shapes(1)` doesn't represent an embedded OLE object, this example fails..


```vb
Worksheets(1).Shapes(1).OLEFormat.Activate
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

