---
title: StartDriver.EffectiveDateDifference Property (Project)
ms.prod: project-server
api_name:
- Project.StartDriver.EffectiveDateDifference
ms.assetid: 9b825839-31de-71f8-9804-015dfd5a293c
ms.date: 06/08/2017
---


# StartDriver.EffectiveDateDifference Property (Project)

Gets the duration between two dates in minutes, using the effective calendar for a manually scheduled task. Read-only  **Long**.


## Syntax

 _expression_. **EffectiveDateDifference**( ** _StartDate_**, ** _FinishDate_** )

 _expression_ An expression that returns a **StartDriver** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartDate_|Required|**Variant**|Arbitrary start date and time, for example, "7/10/2010" or "7/10/2010 2:00:00 PM".|
| _FinishDate_|Required|**Variant**|Arbitrary finish date and time.|

## Remarks

The  **EffectiveDateDifference** property uses the effective calendar for manually scheduled tasks, which allows tasks to start and finish on non-working times. The StartDate and FinishDate arguments can be arbitrary dates. The property and arguments do not affect the task dates.

You can use the  **[EffectiveDateSubtract](startdriver-effectivedatesubtract-property-project.md)**, **[EffectiveDateAdd](startdriver-effectivedateadd-property-project.md)**, and **EffectiveDateDifference** properties to calculate start and finish dates for manually scheduled tasks.

To calculate the date difference for an automatically scheduled task, where you can also specify the calendar, use the  **[DateDifference](application-datedifference-method-project.md)** method.


## Example

The following statement returns the value 480, which shows that the finish date is 8 hours of working time after the start date. 


```vb
Debug.Print ActiveProject.Tasks(3).StartDriver.EffectiveDateDifference("7/1/2009 3:00:00 PM", "7/2/2009 3:00:00 PM")
```

The following statement returns the value -840, which shows that the finish date is 14 hours of working time before the start date. 




```vb
Debug.Print ActiveProject.Tasks(3).StartDriver.EffectiveDateDifference("7/1/2009 3:00:00 PM", "6/30/2009 8:00:00 AM")
```


