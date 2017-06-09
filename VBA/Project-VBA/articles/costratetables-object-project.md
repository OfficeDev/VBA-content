---
title: CostRateTables Object (Project)
ms.prod: project-server
ms.assetid: f08a0a0c-d7ef-f315-5435-804897d5158a
ms.date: 06/08/2017
---


# CostRateTables Object (Project)

 Contains a collection of **[CostRateTable](costratetable-object-project.md)** objects.
 


## Example

 **Using the CostRateTables Collection**
 

 
Use the  **[CostRateTables](resource-costratetables-property-project.md)** property to return a **CostRateTables** collection. The following example lists the standard pay rates for all the cost rate tables of the resource in the active cell.
 

 



```
Dim CRT As CostRateTable, PR As PayRate 

Dim Rates As String 

 

For Each CRT In ActiveCell.Resource.CostRateTables 

 For Each PR In CRT.PayRates 

 Rates = Rates &amp; "CostRateTable " &amp; CRT.Name &amp; ": " &amp; PR.StandardRate &amp; vbCrLf 

 Next PR 

Next CRT 

 

MsgBox Rates
```


## Properties



|**Name**|
|:-----|
|[Application](costratetables-application-property-project.md)|
|[Count](costratetables-count-property-project.md)|
|[Item](costratetables-item-property-project.md)|
|[Parent](costratetables-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
