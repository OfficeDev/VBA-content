---
title: PayRate.CostPerUse Property (Project)
ms.prod: project-server
api_name:
- Project.PayRate.CostPerUse
ms.assetid: 7925d309-afb9-a0f8-7d40-9c2388fdaa1d
ms.date: 06/08/2017
---


# PayRate.CostPerUse Property (Project)

Gets or sets the cost per use for the pay rate. Read/write  **Variant**.


## Syntax

 _expression_. **CostPerUse**

 _expression_ A variable that represents a **PayRate** object.


## Remarks

If the specified pay rate does not exist, the  **CostPerUse** property results in a run-time error 1101.


## Example

The following example prints the cost per use specified in the first pay rate of the first resource for the active project.


```vb
Debug.Print ActiveProject.Resources(1).PayRates(1).CostPerUse
```


