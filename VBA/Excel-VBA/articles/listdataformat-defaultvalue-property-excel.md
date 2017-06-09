---
title: ListDataFormat.DefaultValue Property (Excel)
keywords: vbaxl10.chm758076
f1_keywords:
- vbaxl10.chm758076
ms.prod: excel
api_name:
- Excel.ListDataFormat.DefaultValue
ms.assetid: 503de2f7-878e-a946-9138-10922082bc0d
ms.date: 06/08/2017
---


# ListDataFormat.DefaultValue Property (Excel)

 Returns **Variant** representing the default data type value for a new row in a column. The **Nothing** object is returned when the schema does not specify a default value. Read-only **Variant** .


## Syntax

 _expression_ . **DefaultValue**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

This property is used only for tables linked to a Microsoft SharePoint Foundation site.

In Excel, you cannot set any of the properties associated with the  **[ListDataFormat](listdataformat-object-excel.md)** object. You can set these properties, however, by modifying the list on the SharePoint site.


## Example

The following example displays the setting of the  **DefaultValue** property for the third column of the table in Sheet1 of the active workbook. This code requires a list linked to a SharePoint site.


```vb
Sub ShowDefaultSetting() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 If IsNull(objListCol.ListDataFormat.DefaultValue) Then 
 MsgBox "No default value specified." 
 Else 
 MsgBox objListCol.ListDataFormat.DefaultValue 
 End If 
End Sub
```


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

