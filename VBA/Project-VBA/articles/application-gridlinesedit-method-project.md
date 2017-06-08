---
title: Application.GridlinesEdit Method (Project)
keywords: vbapj.chm2061
f1_keywords:
- vbapj.chm2061
ms.prod: project-server
api_name:
- Project.Application.GridlinesEdit
ms.assetid: 75b9d660-88b5-da71-faf8-215abce897d2
ms.date: 06/08/2017
---


# Application.GridlinesEdit Method (Project)

Edits gridlines.


## Syntax

 _expression_. **GridlinesEdit**( ** _Item_**, ** _NormalType_**, ** _NormalColor_**, ** _Interval_**, ** _IntervalType_**, ** _IntervalColor_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**Integer**|The gridline to edit. Can be one of the following  **[PjGridline](pjgridline-enumeration-project.md)** constants:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>If the Gantt Chart is active: <b>pjBarRows</b> , <b>pjGanttCurrentDate</b> , <b>pjGanttPageBreaks</b> , <b>pjGanttProjectFinish</b> , <b>pjGanttProjectStart</b> , <b>pjGanttRows</b> , <b>pjGanttSheetColumns</b> , <b>pjGanttSheetRows</b> , <b>pjGanttStatusDate</b> , <b>pjGanttTitleHorizontal</b> , <b>pjGanttTitleVertical</b> ,  <b>pjMajorColumns</b> , or <b>pjMinorColumns</b> .</p></li><li><p>If the Calendar view is active: <b>pjCalendarDays</b> , <b>pjCalendarWeeks</b> , <b>pjTitleHorizontal</b> , <b>pjTitleVertical</b> , <b>pjDateBoxTop</b> , or <b>pjDateBoxBottom</b> . 
</p></li><li><p>If the Resource Graph is active: <b>pjMajorVertical</b> , <b>pjMinorVertical</b> , <b>pjHorizontal</b> , <b>pjGraphCurrentDate</b> , <b>pjGraphTitleHorizontal</b> , <b>pjGraphTitleVertical</b> , <b>pjGraphProjectStart</b> , <b>pjGraphProjectFinish</b> , or <b>pjGraphStatusDate</b> . 
</p></li><li><p>If the Task Sheet or Resource Sheet is active: <b>pjSheetColumns</b> , <b>pjSheetRows</b> , <b>pjSheetTitleHorizontal</b> , <b>pjSheetTitleVertical</b> , or <b>pjSheetPageBreaks</b> .</p></li><li><p>If the Task Usage or Resource Usage view is active: <b>pjUsageColumns</b> , <b>pjUsageRows</b> , <b>pjUsageSheetRows</b> , <b>pjUsageSheetColumns</b> , <b>pjUsageTitleHorizontal</b> , <b>pjUsageTitleVertical</b> , or <b>pjUsagePageBreaks</b> .</p></li></ul>|
| _NormalType_|Optional|**Integer**| The type for normal gridlines. Can be one of the following **[PjLineType](pjlinetype-enumeration-project.md)** constants: **pjNoLines**, **pjContinuous**, **pjCloseDot**, **pjDot**, or **pjDash**.|
| _NormalColor_|Optional|**Integer**|The color of normal gridlines. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _Interval_|Optional|**Integer**|A number from 0 to 99 that specifies the interval between gridlines.|
| _IntervalType_|Optional|**Integer**|The type for secondary gridlines. Can be one of the  **[PjLineType](pjlinetype-enumeration-project.md)** constants.|
| _IntervalColor_|Optional|**Integer**|The color of secondary gridlines. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

To edit gridlines where colors can be hexadecimal RGB values, use the  **[GridLinesEditEx](application-gridlineseditex-method-project.md)** method.


## Example

The following example changes the major gridlines to red.


```vb
Sub Gridlines_Edit()    
    'Activate Gantt Chart view 
    ViewApply Name:="&;Gantt Chart" 
    GridlinesEdit Item:=pjMajorColumns, NormalColor:=pjRed 
End Sub
```


