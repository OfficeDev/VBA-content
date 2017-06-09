---
title: StartDriver.EffectiveDateSubtract Property (Project)
ms.prod: project-server
api_name:
- Project.StartDriver.EffectiveDateSubtract
ms.assetid: 14529bd1-9029-d1bc-60a0-b7863cba4d6d
ms.date: 06/08/2017
---


# StartDriver.EffectiveDateSubtract Property (Project)

Gets the date and time that precedes another date by a specified duration, using the effective calendar for a manually scheduled task. Read-only  **Variant**.


## Syntax

 _expression_. **EffectiveDateSubtract**( ** _Date_**, ** _Duration_** )

 _expression_ An expression that returns a **StartDriver** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Date_|Required|**Variant**|Arbitrary date and time, for example, "7/10/2010" or "7/10/2010 2:00:00 PM".|
| _Duration_|Required|**Variant**|Duration to subtract, for example, "3d" or "2w".|

## Remarks

The  **EffectiveDateSubtract** property uses the effective calendar for manually scheduled tasks, which allows tasks to start and finish on non-working times. The property and arguments have no effect on actual task dates.

You can use the  **EffectiveDateSubtract**, **[EffectiveDateAdd](startdriver-effectivedateadd-property-project.md)**, and **[EffectiveDateDifference](startdriver-effectivedatedifference-property-project.md)** properties to calculate start and finish dates for manually scheduled tasks.

To calculate a date for an automatically scheduled task, where you can also specify the calendar, use the  **[DateSubtract](application-datesubtract-method-project.md)** method.


## Example

The following statement returns the value "6/24/2009 8:00:00 AM", which is six days before the specified date. 


```vb
Debug.Print ActiveProject.Tasks(3).StartDriver.EffectiveDateSubtract("7/2/2009", "6d")
```


