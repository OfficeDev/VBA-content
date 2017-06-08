---
title: Application.DateAdd Method (Project)
ms.prod: project-server
api_name:
- Project.Application.DateAdd
ms.assetid: df0da054-495c-c224-ebc8-b47acb78e2af
ms.date: 06/08/2017
---


# Application.DateAdd Method (Project)

Returns the date and time that follows another date by a specified duration, for an automatically scheduled task.


## Syntax

 _expression_. **DateAdd**( ** _StartDate_**, ** _Duration_**, ** _Calendar_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartDate_|Required|**Variant**|The original date to which the duration is added.|
| _Duration_|Required|**Variant**|The duration to add to the start date.|
| _Calendar_|Optional|**Object**|A resource, task, or base calendar object. The default value is the calendar of the active project.|

### Return Value

 **Variant**


## Remarks

To to add a duration to a date for a manually scheduled task, which uses an effective calendar that can include non-working time, use the  **[EffectiveDateAdd](startdriver-effectivedateadd-property-project.md)** property.


## Example

The following example displays the finish date of a three-day automatically scheduled task that begins on 7/11/07 at 8 A.M.


```vb
Sub FindFinishDate() 
 MsgBox Application.DateAdd(StartDate:="7/11/07 8:00 AM", Duration:="3d") 
End Sub
```


