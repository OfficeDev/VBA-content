---
title: OLEObject.Object Property (Excel)
keywords: vbaxl10.chm417076
f1_keywords:
- vbaxl10.chm417076
ms.prod: excel
api_name:
- Excel.OLEObject.Object
ms.assetid: f49881b7-a793-8431-e50d-d56282004699
ms.date: 06/08/2017
---


# OLEObject.Object Property (Excel)

Returns the OLE Automation object associated with this OLE object. Read-only  **Object** .


## Syntax

 _expression_ . **Object**

 _expression_ A variable that represents an **OLEObject** object.


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


[OLEObject Object](oleobject-object-excel.md)

