---
title: CalculatedMember.IsValid Property (Excel)
keywords: vbaxl10.chm686077
f1_keywords:
- vbaxl10.chm686077
ms.prod: excel
api_name:
- Excel.CalculatedMember.IsValid
ms.assetid: 9b0f78c6-3435-6539-aff0-165810668dde
ms.date: 06/08/2017
---


# CalculatedMember.IsValid Property (Excel)

Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session.


## Syntax

 _expression_ . **IsValid**

 _expression_ A variable that represents a **CalculatedMember** object.


## Remarks

This property returns  **True** even if the PivotTable is not connected to its data source. Make sure that the PivotTable is connected before querying the value of the **IsValid** property.


## Example

This example notifies the user about whether the calculated member is valid or not. It assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckValidity() 
 
 Dim pvtTable As PivotTable 
 Dim pvtCache As PivotCache 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Make connection for PivotTable before testing IsValid property. 
 pvtCache.MakeConnection 
 
 ' Check if calculated member is valid. 
 If pvtTable.CalculatedMembers.Item(1).IsValid = True Then 
 MsgBox "The calculated member is valid." 
 Else 
 MsgBox "The calculated member is not valid." 
 End If 
 
End Sub
```


## See also


#### Concepts


[CalculatedMember Object](calculatedmember-object-excel.md)

