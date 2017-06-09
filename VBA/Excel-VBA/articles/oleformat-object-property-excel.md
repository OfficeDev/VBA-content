---
title: OLEFormat.Object Property (Excel)
keywords: vbaxl10.chm632074
f1_keywords:
- vbaxl10.chm632074
ms.prod: excel
api_name:
- Excel.OLEFormat.Object
ms.assetid: be4b7180-34f5-6577-4cfa-b8df017f307a
ms.date: 06/08/2017
---


# OLEFormat.Object Property (Excel)

Returns the OLE Automation object associated with this OLE object. Read-only  **Object** .


## Syntax

 _expression_ . **Object**

 _expression_ A variable that represents an **OLEFormat** object.


## Example

This example inserts text at the beginning of an embedded Word document object on Sheet1. Note that the three statements in the  **With** control structure are WordBasic statements.


```vb
Set wordObj = Worksheets("Sheet1").OLEObjects(1) 
wordObj.Activate 
With wordObj.Object.Application.WordBasic 
 .StartOfDocument 
 .Insert "This is the beginning" 
 .InsertPara 
End With
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-excel.md)

