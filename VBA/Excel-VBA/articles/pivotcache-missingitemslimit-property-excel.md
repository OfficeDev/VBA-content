---
title: PivotCache.MissingItemsLimit Property (Excel)
keywords: vbaxl10.chm227102
f1_keywords:
- vbaxl10.chm227102
ms.prod: excel
api_name:
- Excel.PivotCache.MissingItemsLimit
ms.assetid: ff15a86c-b57f-ed55-bbfa-74e1c5ce753c
ms.date: 06/08/2017
---


# PivotCache.MissingItemsLimit Property (Excel)

Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. Read/write  **[XlPivotTableMissingItems](xlpivottablemissingitems-enumeration-excel.md)** .


## Syntax

 _expression_ . **MissingItemsLimit**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks



| **XlPivotTableMissingItems** can be one of these **XlPivotTableMissingItems** constants.|
| **xlMissingItemsDefault** The default number of unique items per PivotField allowed.|
| **xlMissingItemsMax** The maximum number of unique items per PivotField allowed (32,500).|
| **xlMissingItemsNone** No unique items per PivotField allowed (zero).|
This property can be set to a value between 0 and 32500. If an integer less than zero is specified, this is equivalent to specifying  **xlMissingItemsDefault** . Integers greater than 32,500 can be specified but will have the same effect as specifying **xlMissingItemsMax** .

The  **MissingItemsLimit** property only works for non-OLAP PivotTables; otherwise, a run-time error can occur.


## Example

This example determines the maximum quantity of unique items per field and notifies the user. The example assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckMissingItemsList() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Determine the maximum number of unique items allowed per PivotField and notify the user. 
 Select Case pvtCache.MissingItemsLimit 
 Case xlMissingItemsDefault 
 MsgBox "The default value of unique items per PivotField is allowed." 
 Case xlMissingItemsMax 
 MsgBox "The maximum value of unique items per PivotField is allowed." 
 Case xlMissingItemsNone 
 MsgBox "No unique items per PivotField are allowed." 
 End Select 
 
End Sub
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

