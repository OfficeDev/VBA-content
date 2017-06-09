---
title: ListDataFormat.ReadOnly Property (Excel)
keywords: vbaxl10.chm758084
f1_keywords:
- vbaxl10.chm758084
ms.prod: excel
api_name:
- Excel.ListDataFormat.ReadOnly
ms.assetid: 978c9dc4-fc82-fb26-11f0-e333e43393b6
ms.date: 06/08/2017
---


# ListDataFormat.ReadOnly Property (Excel)

 Returns **True** if the object has been opened as read-only. Read-only **Boolean** .


## Syntax

 _expression_ . **ReadOnly**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

This property is used only for tables that are linked to a SharePoint site.


## Example

The following example displays the setting of the  **ReadOnly** property for the third column of a table in Sheet1 of the active workbook.


```vb
 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.ReadOnly
```


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

