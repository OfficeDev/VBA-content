---
title: ListDataFormat.MaxNumber Property (Excel)
keywords: vbaxl10.chm758080
f1_keywords:
- vbaxl10.chm758080
ms.prod: excel
api_name:
- Excel.ListDataFormat.MaxNumber
ms.assetid: 61262a29-7a35-e351-71fa-0b217285e2b3
ms.date: 06/08/2017
---


# ListDataFormat.MaxNumber Property (Excel)

Returns a  **Variant** containing the maximum value allowed in this field in the list column. Read-only **Variant** .


## Syntax

 _expression_ . **MaxNumber**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

The  **Nothing** object is returned if a maximum value number has not been specified or if the **Type** property setting is such that a maximum value for the column is not applicable.

This property is used only for lists that are linked to a SharePoint site.

In Microsoft Excel, you cannot set any of the properties associated with the  **ListDataFormat** object. You can set these properties, however, by modifying the list on the SharePoint site.


## Example

The following example displays the setting of the  **MaxNumber** property for the third column of a list in Sheet1 of the active workbook.


```vb
 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.MaxNumber
```


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

