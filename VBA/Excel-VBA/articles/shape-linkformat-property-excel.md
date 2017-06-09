---
title: Shape.LinkFormat Property (Excel)
keywords: vbaxl10.chm636129
f1_keywords:
- vbaxl10.chm636129
ms.prod: excel
api_name:
- Excel.Shape.LinkFormat
ms.assetid: f364d08e-aafd-1555-34ee-f0682cde7e19
ms.date: 06/08/2017
---


# Shape.LinkFormat Property (Excel)

Returns a  **[LinkFormat](linkformat-object-excel.md)** object that contains linked OLE object properties. Read-only.


## Syntax

 _expression_ . **LinkFormat**

 _expression_ A variable that represents a **Shape** object.


## Example

This example updates all linked OLE objects on worksheet one.


```vb
For Each s In Worksheets(1).Shapes 
 If s.Type = msoLinkedOLEObject Then s.LinkFormat.Update 
Next
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

