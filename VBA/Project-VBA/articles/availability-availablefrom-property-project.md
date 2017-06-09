---
title: Availability.AvailableFrom Property (Project)
ms.prod: project-server
api_name:
- Project.Availability.AvailableFrom
ms.assetid: 114a1c41-8866-f479-ef08-e099cf7a9968
ms.date: 06/08/2017
---


# Availability.AvailableFrom Property (Project)

Gets the earliest date that a resource is available for work on the project, as specified in the  **Availability** row of the **Resource Availability** grid for the resource. Read/write **Variant**.


## Syntax

 _expression_. **AvailableFrom**

 _expression_ A variable that represents an **Availability** object.


## Remarks

You can set availability information on the  **General** tab of the **Resource Information** dialog box, in the **Resource Availability** grid.

The  **AvailableFrom** property does not return any meaningful information for material resources.


## Example

The following line of code prints the earliest date that the resource is available, during the second availability period. If the availability period is not defined, the code results in run-time error 1101, "The argument value is not valid."


```vb
Debug.Print ActiveProject.Resources(1).Availabilities(2).AvailableFrom
```


