---
title: Resource.AvailableTo Property (Project)
keywords: vbapj.chm131413
f1_keywords:
- vbapj.chm131413
ms.prod: project-server
api_name:
- Project.Resource.AvailableTo
ms.assetid: 27671dd6-77c8-0dea-eed5-943237c10dcd
ms.date: 06/08/2017
---


# Resource.AvailableTo Property (Project)

Gets the latest date that a resource is available for work on the project, as specified in the current row of the **Resource Availability** grid for the resource. Read/write **Variant**.


## Syntax

 _expression_. **AvailableTo**

 _expression_ A variable that represents a **Resource** object.


## Remarks

You can set availability information on the  **General** tab of the **Resource Information** dialog box in the **Resource Availability** grid. The current row is the row in which the date range specified by the **Available From** and **Available To** columns includes the current date.

The  **AvailableTo** property does not return any meaningful information for material resources.


## Example

The following line of code prints the latest date that the resource is available, as of the current date. If the availability period for the current date is not defined, the code prints the date before the next nearest  **Available From** date. If no subsequent availability dates are defined, the code prints "NA".


```vb
Debug.Print ActiveProject.Resources(1).AvailableTo
```


