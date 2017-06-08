---
title: CalculatedMember.SolveOrder Property (Excel)
keywords: vbaxl10.chm686076
f1_keywords:
- vbaxl10.chm686076
ms.prod: excel
api_name:
- Excel.CalculatedMember.SolveOrder
ms.assetid: 45e461ac-4900-000b-cb72-4022bcc1a7c9
ms.date: 06/08/2017
---


# CalculatedMember.SolveOrder Property (Excel)

Returns a  **Long** specifying the value of the calculated member's solve order MDX (Mulitdimensional Expression) argument. The default value is zero. Read-only.


## Syntax

 _expression_ . **SolveOrder**

 _expression_ A variable that represents a **CalculatedMember** object.


## Example

This example determines the solve order for a calculated member and notifies the user. The example assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckSolveOrder() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine solve order and notify user. 
 If pvtTable.CalculatedMembers.Item(1).SolveOrder = 0 Then 
 MsgBox "The solve order is set to the default value." 
 Else 
 MsgBox "The solve order is not set to the default value." 
 End If 
 
End Sub
```


## See also


#### Concepts


[CalculatedMember Object](calculatedmember-object-excel.md)

