---
title: Application.TimescaleEdit Method (Project)
keywords: vbapj.chm902
f1_keywords:
- vbapj.chm902
ms.prod: project-server
api_name:
- Project.Application.TimescaleEdit
ms.assetid: 7f1ee80d-8de3-ebde-9961-105a31c62653
ms.date: 06/08/2017
---


# Application.TimescaleEdit Method (Project)

Enables changing the scale and format of a timescale in a Gantt chart or other timephased view.


## Syntax

 _expression_. **TimescaleEdit**( ** _MajorUnits_**, ** _MinorUnits_**, ** _MajorLabel_**, ** _MinorLabel_**, ** _MajorAlign_**, ** _MinorAlign_**, ** _MajorCount_**, ** _MinorCount_**, ** _MajorTicks_**, ** _MinorTicks_**, ** _Enlarge_**, ** _Separator_**, ** _MajorUseFY_**, ** _MinorUseFY_**, ** _TopUnits_**, ** _TopLabel_**, ** _TopAlign_**, ** _TopCount_**, ** _TopTicks_**, ** _TopUseFY_**, ** _TierCount_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MajorUnits_|Optional|**Variant**|Time units for the middle timescale tier. Specify with the  **[PjTimescaleUnit](pjtimescaleunit-enumeration-project.md)** enumeration. The default value is **pjTimescaleWeeks** (3).|
| _MinorUnits_|Optional|**Variant**|Time units for the bottom timescale tier. Specify with the  **[PjTimescaleUnit](pjtimescaleunit-enumeration-project.md)** enumeration. The default value is **pjTimescaleDays** (2).|
| _MajorLabel_|Optional|**Variant**|Date or time labels for the middle timescale tier. Specify with one of the following enumerations:  **[PjDateLabel](pjdatelabel-enumeration-project.md)**, **[PjDayLabel](pjdaylabel-enumeration-project.md)**, **[PjMonthLabel](pjmonthlabel-enumeration-project.md)**, or **[PjWeekLabel](pjweeklabel-enumeration-project.md)**. The default value is **pjWeekLabelWeek_mmm_dd_yyy** (13). For example, default values are **Mar 7, '10** and **Mar 14, '10**.|
| _MinorLabel_|Optional|**Variant**|Date or time labels for the bottom timescale tier. Specify with one of the following enumerations:  **[PjDateLabel](pjdatelabel-enumeration-project.md)**, **[PjDayLabel](pjdaylabel-enumeration-project.md)**, **[PjMonthLabel](pjmonthlabel-enumeration-project.md)**, or **[PjWeekLabel](pjweeklabel-enumeration-project.md)**. The default value is **pjDayLabelDay_di** (20). For example, default values are **S**,  **M**, and  **T**.|
| _MajorAlign_|Optional|**Variant**|The  **[PjAlignment](pjalignment-enumeration-project.md)** enumeration specifies how to align labels within each time period of the middle tier. The default is value is **pjLeft** (0).|
| _MinorAlign_|Optional|**Variant**|The  **[PjAlignment](pjalignment-enumeration-project.md)** enumeration specifies how to align labels within each time period of the bottom tier. The default is value is **pjLeft** (0).|
| _MajorCount_|Optional|**Variant**|Specifies the time unit interval in which to show labels for the middle tier. For example, if the time unit is weeks, a value of 1 shows a label every week; a value of 2 shows a label every two weeks.|
| _MinorCount_|Optional|**Variant**|Specifies the time unit interval in which to show labels for the bottom tier. For example, if the time unit is hours, a value of 1 shows a label every hour; a value of 2 shows a label every two hours.|
| _MajorTicks_|Optional|**Variant**|Specifies whether to show tick marks that separate time periods in the middle tier. For example, if the time unit is days, a value of  **False** removes the tick marks between days.|
| _MinorTicks_|Optional|**Variant**|Specifies whether to show tick marks that separate time periods in the bottom tier. For example, if the time unit is hours, a value of  **False** removes the tick marks between hours.|
| _Enlarge_|Optional|**Variant**|Specifies the percent of horizontal expansion of the timescale. For example, a value of 150 expands the timescale 150%.|
| _Separator_|Optional|**Variant**|Specifies whether to show the lines that separate the top, middle, and bottom tiers of the timescale. For example, a value of  **False** removes the lines.|
| _MajorUseFY_|Optional|**Variant**|Specifies whether to base the middle tier labels on the fiscal year. The default value is  **False**.|
| _MinorUseFY_|Optional|**Variant**|Specifies whether to base the bottom tier labels on the fiscal year. The default value is  **False**.|
| _TopUnits_|Optional|**Variant**|Time units for the top timescale tier. Specify with the  **[PjTimescaleUnit](pjtimescaleunit-enumeration-project.md)** enumeration. The default value is **pjTimescaleMonths** (2).|
| _TopLabel_|Optional|**Variant**|Date or time labels for the top timescale tier. Specify with one of the following enumerations:  **[PjDateLabel](pjdatelabel-enumeration-project.md)**, **[PjDayLabel](pjdaylabel-enumeration-project.md)**, **[PjMonthLabel](pjmonthlabel-enumeration-project.md)**, or **[PjWeekLabel](pjweeklabel-enumeration-project.md)**. The default value is **pjDayLabelDay_di** (20). For example, default values are **S**,  **M**, and  **T**.|
| _TopAlign_|Optional|**Variant**|The  **[PjAlignment](pjalignment-enumeration-project.md)** enumeration specifies how to align labels within each time period of the top tier. The default is value is **pjLeft** (0).|
| _TopCount_|Optional|**Variant**|Specifies the time unit interval in which to show labels for the top tier. For example, if the time unit is months, a value of 1 shows a label every month; a value of 2 shows a label every two months.|
| _TopTicks_|Optional|**Variant**|Specifies whether to show tick marks that separate time periods in the top tier. For example, if the time unit is months, a value of  **False** removes the tick marks between months.|
| _TopUseFY_|Optional|**Variant**|Specifies whether to base the top tier labels on the fiscal year. The default value is  **False**.|
| _TierCount_|Optional|**Variant**|Specifies the number of timescale tiers. The integer value 3 shows all three tiers; the value 2 is default and shows the middle and bottom tiers; the value 1 shows only the middle tier.|

### Return Value

 **Boolean**


## Remarks

To manually edit a timescale in Project, right-click the timescale, and then choose  **Timescale**. Executing the  **TimescaleEdit** method with no parameters displays the **Timescale** dialog box. If the user choose **Cancel**,  **TimescaleEdit** returns **False**. If the user makes valid changes and chooses **OK**,  **TimescaleEdit** returns **True**.


## Example

The following example sets the timescale to three tiers, where the top tier units are months, the top labels are the month name and year, the middle tier units are weeks, and the middle tier labels are the month and day numbers. For example, top tier labels are  **May 2012** and **June 2012**, and middle tier labels are  **5/27** and **6/3**.


```
TimescaleEdit TierCount:=3, _ 
    TopUnits:=PjTimescaleUnit.pjTimescaleMonths, _ 
    TopLabel:=PjMonthLabel.pjMonthLabelMonth_mmmm_yyyy, _ 
    MajorUnits:=PjTimescaleUnit.pjTimescaleWeeks, _ 
    MajorLabel:=PjWeekLabel.pjWeekLabelWeek_mm_dd
```


 **Note**  Values for the label time range in the  _TopLabel_, _MajorLabel_, and _MinorLabel_ parameters must be compatible with the time unit of the specified timescale tier. For example, if the time unit of the bottom tier is hours, the parameter value `MinorLabel:=PjDateLabel.pjHour_hhAM` is valid. However, the parameter value `MinorLabel:=PjDateLabel.pjHalfYear_hhh_Half` results in the run time error 1101: "The argument value is not valid."


