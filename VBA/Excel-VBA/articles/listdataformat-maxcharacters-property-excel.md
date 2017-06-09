---
title: ListDataFormat.MaxCharacters Property (Excel)
keywords: vbaxl10.chm758079
f1_keywords:
- vbaxl10.chm758079
ms.prod: excel
api_name:
- Excel.ListDataFormat.MaxCharacters
ms.assetid: b8d73844-6f2b-7888-8268-a27cbfcc709c
ms.date: 06/08/2017
---


# ListDataFormat.MaxCharacters Property (Excel)

Returns a  **Long** containing the maximum number of characters allowed in the **[ListColumn](listcolumn-object-excel.md)** object if the **[Type](listdataformat-type-property-excel.md)** property is set to **xlListDataTypeText** or **xlListDataTypeMultiLineText** . Read-only **Long** .


## Syntax

 _expression_ . **MaxCharacters**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

Returns -1 for columns whose  **Type** property is set to a non-text value.

This property is used only for lists that are linked to a SharePoint site.

In Microsoft Excel, you cannot set any of the properties associated with the  **ListDataFormat** object. You can set these properties, however, by modifying the list on the SharePoint site.


## Example

The following example displays the setting of the  **MaxCharacters** property for the third column of a list in Sheet1 of the active workbook.


```vb
 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.MaxCharacters
```


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

