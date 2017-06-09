---
title: Application.CalendarDateBoxes Method (Project)
keywords: vbapj.chm2340
f1_keywords:
- vbapj.chm2340
ms.prod: project-server
api_name:
- Project.Application.CalendarDateBoxes
ms.assetid: 3870fa41-ef58-8b5d-efe1-b8b3d3a03835
ms.date: 06/08/2017
---


# Application.CalendarDateBoxes Method (Project)

Customizes the top and bottom bands of date boxes in the Calendar view.


## Syntax

 _expression_. **CalendarDateBoxes**( ** _TopLeft_**, ** _TopRight_**, ** _BottomLeft_**, ** _BottomRight_**, ** _TopColor_**, ** _BottomColor_**, ** _TopPattern_**, ** _BottomPattern_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TopLeft_|Optional|**Long**|The format for dates in the upper-left corner of each date box. Can be one of the  **[PjCalendarDateLabel](pjcalendardatelabel-enumeration-project.md)** constants.|
| _TopRight_|Optional|**Long**|The format for dates in the upper-right corner of each date box. Can be one of the  **PjCalendarDateLabel** constants.|
| _BottomLeft_|Optional|**Long**|The format for dates in the lower-left corner of each date box. Can be one of the  **PjCalendarDateLabel** constants.|
| _BottomRight_|Optional|**Long**|The format for dates in the lower-right corner of each date box. Can be one of the  **PjCalendarDateLabel** constants.|
| _TopColor_|Optional|**Long**|The color of the top band in each date box. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _BottomColor_|Optional|**Long**|The color of the bottom band in each date box. Can be one of the  **PjColor** constants.|
| _TopPattern_|Optional|**Long**|The pattern of the top band in each date box. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _BottomPattern_|Optional|**Long**|The pattern of the bottom band in each date box. Can be one of the  **PjFillPattern** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **CalendarDateBoxes** method with no arguments displays the **Timescale** dialog box with the **Date Boxes** tab selected.

To edit calendar date boxes where the colors can be RGB values, use the  **[CalendarDateBoxesEx](application-calendardateboxesex-method-project.md)** method.


## Example

The following example displays the day of the week (for example, Thursday) in the upper-left corner, the month and date (for example, Jan 31) in the upper-right corner, the day of the year and year (for example, 70 2012) in the bottom-left corner of each date box in the calendar, and sets the background colors of the top band and the bottom band.


```vb
Sub FormatCalendarDays() 
    CalendarDateBoxes Topleft:=pjDay_dddd, TopRight:=pjDay_mmm_dd, _
        BottomLeft:=pjCalendarLabelDayOfYear_dd_yyyy, _ 
        TopColor:=PjColor.pjSilver, BottomColor:=PjColor.pjYellow 
End Sub
```


