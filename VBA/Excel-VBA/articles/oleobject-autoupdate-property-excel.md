---
title: OLEObject.AutoUpdate Property (Excel)
keywords: vbaxl10.chm417075
f1_keywords:
- vbaxl10.chm417075
ms.prod: excel
api_name:
- Excel.OLEObject.AutoUpdate
ms.assetid: 3834c552-a282-ab75-781e-42c055346b7d
ms.date: 06/08/2017
---


# OLEObject.AutoUpdate Property (Excel)

 **True** if the OLE object is updated automatically when the source changes. Valid only if the object is linked (its **OLEType** property must be **xlOLELink** ). Read-only **Boolean** .


## Syntax

 _expression_ . **AutoUpdate**

 _expression_ A variable that represents an **OLEObject** object.


## Example

This example displays the status of automatic updating for all OLE objects on Sheet1.


```vb
Worksheets("Sheet1").Activate 
Range("A1").Value = "Name" 
Range("B1").Value = "Link Status" 
Range("C1").Value = "AutoUpdate Status" 
i = 2 
For Each obj In ActiveSheet.OLEObjects 
 Cells(i, 1) = obj.Name 
 If obj.OLEType = xlOLELink Then 
 Cells(i, 2) = "Linked" 
 Cells(i, 3) = obj.AutoUpdate 
 Else 
 Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)

