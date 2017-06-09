---
title: CostRateTable.PayRates Property (Project)
ms.prod: project-server
api_name:
- Project.CostRateTable.PayRates
ms.assetid: 260d9e77-9fce-5169-687f-027995c73273
ms.date: 06/08/2017
---


# CostRateTable.PayRates Property (Project)

Gets a  **[PayRates](payrate-object-project.md)** collection that represents the various pay rates on the cost rate table for a resource. Read-only **PayRates**.


## Syntax

 _expression_. **PayRates**

 _expression_ A variable that represents a **CostRateTable** object.


## Remarks

For the  **Resource** object, the **PayRates** property returns pay rates for cost rate table A, the default table.


## Example

The following example lists the standard pay rates for all the cost rate tables of the resource in the active cell.


```vb
Sub ListPayRates() 
 Dim CRT As CostRateTable, PR As PayRate 
 Dim Rates As String 
 
 For Each CRT In ActiveCell.Resource.CostRateTables 
 For Each PR In CRT.PayRates 
 Rates = Rates &; "CostRateTable " &; CRT.Name &; ": " &; _ 
 PR.StandardRate &; " (Effective " &; PR.EffectiveDate &; _ 
 ")" &; vbCrLf 
 Next PR 
 Next CRT 
 
 MsgBox Rates 
 
End Sub
```


