---
title: Application.DateDifference Method (Project)
keywords: vbapj.chm131207
f1_keywords:
- vbapj.chm131207
ms.prod: project-server
api_name:
- Project.Application.DateDifference
ms.assetid: 7f34e866-5cd3-971d-42ee-39e7768c1273
ms.date: 06/08/2017
---


# Application.DateDifference Method (Project)

Returns the duration between two dates in minutes, for an automatically scheduled task.


## Syntax

 _expression_. **DateDifference**( ** _StartDate_**, ** _FinishDate_**, ** _Calendar_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartDate_|Required|**Variant**|The date used as the beginning of the duration.|
| _FinishDate_|Required|**Variant**|The date used as the end of the duration.|
| _Calendar_|Optional|**Object**|A resource or task base calendar object. The default value is the calendar of the active project.|

### Return Value

 **Long**


## Remarks

To get a difference between two dates for a manually scheduled task, which uses an effective calendar that can include non-working time, use the  **[EffectiveDateDifference](startdriver-effectivedatedifference-property-project.md)** property.


## Example

The following example displays the duration of a task that begins on 7/11/97 at 8 A.M. and ends on 7/13/97 at 5:00 P.M.


```vb
Sub FindDuration() 
 MsgBox Application.DateDifference ("7/11/97 8:00 AM", "7/13/97 5:00 PM") 
End Sub
```


