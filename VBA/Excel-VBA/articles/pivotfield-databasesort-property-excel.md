---
title: PivotField.DatabaseSort Property (Excel)
keywords: vbaxl10.chm240130
f1_keywords:
- vbaxl10.chm240130
ms.prod: excel
api_name:
- Excel.PivotField.DatabaseSort
ms.assetid: 18c75552-0993-24b6-e31f-7912e69ac933
ms.date: 06/08/2017
---


# PivotField.DatabaseSort Property (Excel)

When set to  **True** , manual repositioning of items in a PivotTable field is allowed. Returns **True** , if the field has no manually positioned items. Read/write **Boolean** .


## Syntax

 _expression_ . **DatabaseSort**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The  **DatabaseSort** property returns **False** if the data source is not an Online Analytical Processing (OLAP) data source.

This property returns  **True** if the data source is OLAP and neither custom ordering nor automatic sorting has been applied to the field.

Setting the  **DatabaseSort** property to **True** , for an OLAP PivotTable, will remove any custom ordering or automatic sort applied to the field (in other words, the PivotTable reverts to the default behavior when the connection was made).

Setting the  **DatabaseSort** property to **False** will cause the sort order to be the current order of the items, if no automatic sort is applied.

Setting the  **DatabaseSort** property to either **True** or **False** causes an Update.

Setting the  **DatabaseSort** property to **True** for a non-OLAP source or an OLAP data field causes a run-time error.


## Example

The following example determines if the data source is an OLAP data source and notifies the user. This example assumes an OLAP PivotTable exists on the active worksheet.


```vb
Sub UseDatabaseSort() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields("[Product].[Product Family]") 
 
 ' Determine source type for the PivotTable report. 
 If pvtField.DatabaseSort = True Then 
 MsgBox "The source is OLAP; you can manually reorder items." 
 Else 
 MsgBox "The data source might not be OLAP." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

