---
title: PivotItemList Object (Excel)
keywords: vbaxl10.chm720072
f1_keywords:
- vbaxl10.chm720072
ms.prod: excel
api_name:
- Excel.PivotItemList
ms.assetid: 2b0fc8e5-6073-9cb1-2217-1e8715cddb1e
ms.date: 06/08/2017
---


# PivotItemList Object (Excel)

A collection of all the  **[PivotItem](pivotitem-object-excel.md)** objects in the specified PivotTable.


## Remarks

 Each **PivotItem** represents an item in a PivotTable field.

Use the  **[RowItems](pivotcell-rowitems-property-excel.md)** or **[ColumnItems](pivotcell-columnitems-property-excel.md)** property of the **[PivotCell](pivotcell-object-excel.md)** object to return a **PivotItemList** collection.


## Example

Once a  **PivotItemList** collection is returned, you can use the **[Item](pivotitems-item-method-excel.md)** method to identify a particular **PivotItem** list. The following example displays the **PivotItem** list associated with cell B5 to the user. This example assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckPivotItemList() 
 
 ' Identify contents associated with PivotItemList. 
 MsgBox "Contents associated with cell B5: " &; _ 
 Application.Range("B5").PivotCell.RowItems.Item(1) 
 
End Sub
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


