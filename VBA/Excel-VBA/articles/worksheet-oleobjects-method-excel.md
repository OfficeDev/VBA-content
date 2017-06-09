---
title: Worksheet.OLEObjects Method (Excel)
keywords: vbaxl10.chm175108
f1_keywords:
- vbaxl10.chm175108
ms.prod: excel
api_name:
- Excel.Worksheet.OLEObjects
ms.assetid: 3f178081-2a42-a751-ae79-8ca149d8ec45
ms.date: 06/08/2017
---


# Worksheet.OLEObjects Method (Excel)

Returns an object that represents either a single OLE object (an  **[OLEObject](oleobject-object-excel.md)** ) or a collection of all OLE objects (an **[OLEObjects](oleobjects-object-excel.md)** collection) on the chart or sheet. Read-only.


## Syntax

 _expression_ . **OLEObjects**( **_Index_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the OLE object.|

### Return Value

Object


## Example

This example creates a list of link types for OLE objects on Sheet1. The list appears on a new worksheet created by the example.


```vb
Set newSheet = Worksheets.Add 
i = 2 
newSheet.Range("A1").Value = "Name" 
newSheet.Range("B1").Value = "Link Type" 
For Each obj In Worksheets("Sheet1").OLEObjects 
 newSheet.Cells(i, 1).Value = obj.Name 
 If obj.OLEType = xlOLELink Then 
 newSheet.Cells(i, 2) = "Linked" 
 Else 
 newSheet.Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

