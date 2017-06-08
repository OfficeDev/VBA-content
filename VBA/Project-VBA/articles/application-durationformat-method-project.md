---
title: Application.DurationFormat Method (Project)
keywords: vbapj.chm131212
f1_keywords:
- vbapj.chm131212
ms.prod: project-server
api_name:
- Project.Application.DurationFormat
ms.assetid: 37970edc-c6f9-66b7-7c0d-b22beb8a36c1
ms.date: 06/08/2017
---


# Application.DurationFormat Method (Project)

Returns a duration in the specified units.


## Syntax

 _expression_. **DurationFormat**( ** _Duration_**, ** _Units_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Duration_|Required|**Variant**|The duration to be expressed.|
| _Units_|Optional|**Long**|The units used to express the duration. Can be one of the  **[PjFormatUnit](pjformatunit-enumeration-project.md)** constants.|

### Return Value

 **String**


## Remarks

The time label that appears next to the duration uses the format specified by the  ** _timescale_ as:** setting on the **Edit** tab of the **Options** dialog box, where ** _timescale_** is "Minutes", "Hours", "Days", "Weeks", "Months", or "Years".

 For example, if _Duration_ is "2w", _Units_ is **pjDays**, and the **Days as:** setting is "day", the **DurationFormat** method returns "10 days".


## Example

The following example displays the duration of the selected task in weeks.


```vb
Sub DurationInWeeks() 
 MsgBox DurationFormat(ActiveCell.Task.Duration, pjWeeks) 
End Sub
```


