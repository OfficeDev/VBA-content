---
title: ListDataFormat.MinNumber Property (Excel)
keywords: vbaxl10.chm758081
f1_keywords:
- vbaxl10.chm758081
ms.prod: excel
api_name:
- Excel.ListDataFormat.MinNumber
ms.assetid: 97d5cc98-ab65-7355-0a54-3b48d7b15cf5
ms.date: 06/08/2017
---


# ListDataFormat.MinNumber Property (Excel)

 Returns a **Variant** containing the minimum value allowed in this field in the list column. This can be a negative floating point number. Read-only **Variant** .


## Syntax

 _expression_ . **MinNumber**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

 This property will return the **Nothing** object if no value has been specified for this field or if the setting of the **Type** property is such that a minimum value is not applicable to the column.

This property is used only for lists that are linked to a SharePoint site.

In Microsoft Excel, you cannot set any of the properties associated with the  **ListDataFormat** object. You can set these properties, however, by modifying the list on the SharePoint site.


## Example

The following example displays the setting of the  **MinNumber** property for the third column of a list in Sheet1 of the active workbook.


```vb
 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.MinNumber
```


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

