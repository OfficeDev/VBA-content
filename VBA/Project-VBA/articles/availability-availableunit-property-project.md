---
title: Availability.AvailableUnit Property (Project)
ms.prod: project-server
api_name:
- Project.Availability.AvailableUnit
ms.assetid: a22d2325-e512-08c5-608f-0fadce9d33e5
ms.date: 06/08/2017
---


# Availability.AvailableUnit Property (Project)

Gets or sets the percentage of time the resource is available during the availability period. Read/write  **Double**.


## Syntax

 _expression_. **AvailableUnit**

 _expression_ A variable that represents an **Availability** object.


## Remarks

If the  **AvailableUnit** value is 100, the resource is available 100% of the time.

You can set availability information on the  **General** tab of the **Resource Information** dialog box, in the **Resource Availability** grid.


## Example

The following line of code prints the percentage of time the resource is available during the second availability period. If the availability period is not defined, the code results in run-time error 1101, "The argument value is not valid."


```vb
Debug.Print ActiveProject.Resources(1).Availabilities(2).AvailableUnit
```


