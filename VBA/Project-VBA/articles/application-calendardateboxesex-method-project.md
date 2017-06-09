---
title: Application.CalendarDateBoxesEx Method (Project)
keywords: vbapj.chm2148
f1_keywords:
- vbapj.chm2148
ms.prod: project-server
api_name:
- Project.Application.CalendarDateBoxesEx
ms.assetid: a6c1fffd-ce21-d3ef-348f-1f41b5231005
ms.date: 06/08/2017
---


# Application.CalendarDateBoxesEx Method (Project)

Customizes the top and bottom bands of date boxes in the Calendar view.


## Syntax

 _expression_. **CalendarDateBoxesEx**( ** _TopLeft_**, ** _TopRight_**, ** _BottomLeft_**, ** _BottomRight_**, ** _TopColor_**, ** _BottomColor_**, ** _TopPattern_**, ** _BottomPattern_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TopLeft_|Optional|**Long**|The format for dates in the upper-left corner of each date box. Can be one of the  **[PjCalendarDateLabel](pjcalendardatelabel-enumeration-project.md)** constants.|
| _TopRight_|Optional|**Long**|The format for dates in the upper-right corner of each date box. Can be one of the  **PjCalendarDateLabel** constants.|
| _BottomLeft_|Optional|**Long**|The format for dates in the lower-left corner of each date box. Can be one of the  **PjCalendarDateLabel** constants.|
| _BottomRight_|Optional|**Long**|The format for dates in the lower-right corner of each date box. Can be one of the  **PjCalendarDateLabel** constants.|
| _TopColor_|Optional|**Long**|The color of the top band in each date box. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|
| _BottomColor_|Optional|**Long**|The color of the bottom band in each date box. Can be a hexadecimal value for the RGB color.|
| _TopPattern_|Optional|**Long**|The pattern of the top band in each date box. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _BottomPattern_|Optional|**Long**|The pattern of the bottom band in each date box. Can be one of the  **PjFillPattern** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **CalendarDateBoxesEx** method without specifying any arguments displays the **Timescale** dialog box with the **Date Boxes** tab selected.


## Example

The following example displays the day of the week (for example, Thursday) in the upper-left corner, the month and date (for example, Jan 31) in the upper-right corner, the day of the year and year (for example, 70 2010) in the bottom-left corner of each date box in the calendar, and sets the background color of the top band to silver and the background color of the bottom band to a light yellow.


```vb
Sub FormatCalendarDays() 
      CalendarDateBoxesEx Topleft:=pjOverflowIndicator, TopRight:=pjDay_mmm_dd, _ 
        BottomLeft:=pjCalendarLabelDayOfYear_dd_yyyy, _ 
        TopColor:=&;HE0E8E8, BottomColor:=&;H1E8E8 
End Sub
```


 **Note**  If you use any of the  **PjColor** enumeration constants for the _TopColor_ or _BottomColor_ parameters, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **CalendarDateBoxesEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[CalendarDateBoxes](application-calendardateboxes-method-project.md)** method.


