---
title: Resource.AvailableFrom Property (Project)
keywords: vbapj.chm131412
f1_keywords:
- vbapj.chm131412
ms.prod: project-server
api_name:
- Project.Resource.AvailableFrom
ms.assetid: a79d0ce3-1c58-25cc-f06a-6c55961b9e0c
ms.date: 06/08/2017
---


# Resource.AvailableFrom Property (Project)

Gets the earliest date that a resource is available for work on the project, as specified in the current row of the **Resource Availability** grid for the resource. Read/write **Variant**.


## Syntax

 _expression_. **AvailableFrom**

 _expression_ A variable that represents a **Resource** object.


## Remarks

You can set availability information on the  **General** tab of the **Resource Information** dialog box, in the **Resource Availability** grid. The current row is the row in which the date range specified by the **Available From** and **Available To** columns includes the current date.

The  **AvailableFrom** property does not return any meaningful information for material resources.


## Example

The following line of code prints the earliest date that the resource is available, as of the current date. If the availability period for the current date is not defined, the code prints the date after the most recent  **Available To** date. If no previous availability dates are defined, the code prints "NA".


```vb
Debug.Print ActiveProject.Resources(1).AvailableFrom
```


