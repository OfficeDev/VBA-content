---
title: Availability.AvailableTo Property (Project)
keywords: vbapj.chm132560
f1_keywords:
- vbapj.chm132560
ms.prod: project-server
api_name:
- Project.Availability.AvailableTo
ms.assetid: fe1b9efc-b981-5dc0-fbaa-a29c098e2628
ms.date: 06/08/2017
---


# Availability.AvailableTo Property (Project)

Gets the latest date that a resource is available for work on the project, as specified in the  **Availability** row of the **Resource Availability** grid for the resource. Read/write **Variant**.


## Syntax

 _expression_. **AvailableTo**

 _expression_ A variable that represents an **Availability** object.


## Remarks

You can set availability information on the  **General** tab of the **Resource Information** dialog box, in the **Resource Availability** grid.

The  **AvailableTo** property does not return any meaningful information for material resources.


## Example

The following line of code prints the latest date that the resource is available, during the second availability period. If the availability period is not defined, the code results in run-time error 1101, "The argument value is not valid."


```vb
Debug.Print ActiveProject.Resources(1).Availabilities(2).AvailableFrom
```


