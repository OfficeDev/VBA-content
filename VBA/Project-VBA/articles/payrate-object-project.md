---
title: PayRate Object (Project)
ms.prod: project-server
api_name:
- Project.PayRate
ms.assetid: 4c8ba1f3-bf18-2179-5f50-c090c63e46b9
ms.date: 06/08/2017
---


# PayRate Object (Project)


 

Represents a line of rates from the cost rate table of a resource. The  **PayRate** object is a member of the **[PayRates](payrates-object-project.md)** collection.
 
 **Using the PayRate Object**
 
Use  **PayRates** (*Index* ), where*Index* is the pay rate index number or date for which to return the rates in effect, to return a single **PayRate** object. The following example returns the standard pay rate for Tamara's first row of rates in cost rate table C.
 



```
ActiveProject.Resources("Tamara").CostRateTables("C").PayRates(1).StandardRate
```

 **Using the PayRates Collection**
 
Use the  **[PayRates](costratetable-payrates-property-project.md)** property to return a **PayRates** collection. The following example lists the standard pay rates for all the cost rate tables of the resource in the active cell.
 



```
Dim CRT As CostRateTable
DIM PR As PayRate
Dim Rates As String

For Each CRT In ActiveCell.Resource.CostRateTables
    For Each PR In CRT.PayRates
        Rates = Rates &amp; "CostRateTable " &amp; CRT.Name &amp; ": " &amp; PR.StandardRate &amp; vbCrLf
    Next PR
Next CRT
    
MsgBox Rates
```

Use the  **[Add](payrates-add-method-project.md)** method to add a **PayRate** object to the **PayRates** collection. The following example adds a line to Tamara's cost rate table "C" with an effective date of September 1, 2012, a standard rate of $40.00 per hour, an overtime rate of $60.00 per hour, and a per-use cost of $0.
 



```
ActiveProject.Resources("Tamara").CostRateTables("C").PayRates.Add "9/1/2012", "$40/h", "$60/h", "$0"
```


## Methods



|**Name**|
|:-----|
|[Delete](payrate-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](payrate-application-property-project.md)|
|[CostPerUse](payrate-costperuse-property-project.md)|
|[EffectiveDate](payrate-effectivedate-property-project.md)|
|[Index](payrate-index-property-project.md)|
|[OvertimeRate](payrate-overtimerate-property-project.md)|
|[Parent](payrate-parent-property-project.md)|
|[StandardRate](payrate-standardrate-property-project.md)|

