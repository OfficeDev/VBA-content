---
title: CostRateTable Object (Project)
ms.prod: project-server
api_name:
- Project.CostRateTable
ms.assetid: ca514e06-3542-00f1-5221-a609378d2392
ms.date: 06/08/2017
---


# CostRateTable Object (Project)


 

Represents a collection of pay rates for a resource. The  **CostRateTable** object is a member of the **[CostRateTables](costratetables-object-project.md)** collection.
 
Use  **CostRateTables(***Index* **)**, where*Index* is the index number or name of the cost rate table, to return a single **CostRateTable** object.
 
 **Using the CostRateTable Object**
 
The following example changes the standard rate on one of a resource's pay rate tables. 
 



```
Dim GovtRates As CostRateTable 
 
Set GovtRates = ActiveProject.Resources("Bob").CostRateTables("B") 
GovtRates.PayRates(1).StandardRate = "$10/h"
```


## Properties



|**Name**|
|:-----|
|[Application](costratetable-application-property-project.md)|
|[Index](costratetable-index-property-project.md)|
|[Name](costratetable-name-property-project.md)|
|[Parent](costratetable-parent-property-project.md)|
|[PayRates](costratetable-payrates-property-project.md)|

