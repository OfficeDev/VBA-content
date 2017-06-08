---
title: CalculatedMembers.Item Property (Excel)
keywords: vbaxl10.chm684074
f1_keywords:
- vbaxl10.chm684074
ms.prod: excel
api_name:
- Excel.CalculatedMembers.Item
ms.assetid: 82ba55c7-0c16-df11-ac32-40868f57d2e1
ms.date: 06/08/2017
---


# CalculatedMembers.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **CalculatedMembers** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

The following example notifies the user if the calculated member is valid or not. This example assumes a PivotTable exists on the active worksheet that contains either a valid or invalid calculated member.


```vb
Sub CheckValidity() 
 
 Dim pvtTable As PivotTable 
 Dim pvtCache As PivotCache 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Handle run-time error if external source is not an OLEDB data source. 
 On Error GoTo Not_OLEDB 
 
 ' Check connection setting and make connection if necessary. 
 If pvtCache.IsConnected = False Then 
 pvtCache.MakeConnection 
 End If 
 
 ' Check if calculated member is valid. 
 If pvtTable.CalculatedMembers.Item(1).IsValid = True Then 
 MsgBox "The calculated member is valid." 
 Else 
 MsgBox "The calculated member is not valid." 
 End If 
 
Not_OLEDB: 
MsgBox "The source is not an OLEDB data source." 
Exit Sub 
 
End Sub
```


## See also


#### Concepts


[CalculatedMembers Collection](calculatedmembers-object-excel.md)

