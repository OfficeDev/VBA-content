---
title: ListDataFormat.IsPercent Property (Excel)
keywords: vbaxl10.chm758077
f1_keywords:
- vbaxl10.chm758077
ms.prod: excel
api_name:
- Excel.ListDataFormat.IsPercent
ms.assetid: 34154cf9-358a-0db9-4b93-fe3b3f2b8dce
ms.date: 06/08/2017
---


# ListDataFormat.IsPercent Property (Excel)

Returns a  **Boolean** value. Returns **True** only if the number data for the **[ListColumn](listcolumn-object-excel.md)** object will be shown in percentage formatting. Read-only **Boolean** . Read-only.


## Syntax

 _expression_ . **IsPercent**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

This property is used only for lists that are linked to a Microsoft SharePoint Foundation site.

In Excel, you cannot set any of the properties associated with the  **ListDataFormat** object. You can set these properties, however, by modifying the list on the SharePoint site.


## Example

The following example returns the setting of the  **IsPercent** property for the third column of the list in Sheet1 of the active workbook.


```vb
Function GetIsPercent() As Boolean 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 GetIsPercent = objListCol.ListDataFormat.IsPercent 
End Function
```


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

