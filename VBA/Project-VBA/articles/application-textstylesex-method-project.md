---
title: Application.TextStylesEx Method (Project)
keywords: vbapj.chm901
f1_keywords:
- vbapj.chm901
ms.prod: project-server
api_name:
- Project.Application.TextStylesEx
ms.assetid: 674c16c8-8ba5-604f-494c-3b59017e1207
ms.date: 06/08/2017
---


# Application.TextStylesEx Method (Project)

Sets the text styles for tasks and resources in the active view.

## Syntax

_expression_. **TextStylesEx** (**_Item_**, **_Font_**, **_Size_**, **_Bold_**, **_Italic_**, **_Underline_**, **_Color_**, **_CellColor_**, **_Pattern_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Integer**|The type of text to change. Can be one of the following **[PjTextItem](pjtextitem-enumeration-project.md)** constants. See [If the Gantt Chart is active](#if-the-gantt-chart-is-active).|
| _Font_|Optional|**String**|The name of the font. The _Font_ argument is ignored if the active view is the **Network Diagram** and _Item_ is not **pjAll**.|
| _Size_|Optional|**Integer**|The size of the font in points. The _Size_ argument is ignored if the active view is the **Network Diagram** and _Item_ is not **pjAll**.|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the font. Can be one of the **[PjColor](pjcolor-enumeration-project.md)** constants (see the [PjColor constants](#pjcolor-constants) table).|
| _CellColor_|Optional|**Long**|The background color of the cell. Can be one of the **PjColor** constants.|
| _Pattern_|Optional|**Integer**|The background pattern of the cell. Can be one of the **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|

<br/>

#### If the Gantt Chart is active

|||
|:-----|:-----|
|**pjAll**|**pjGanttMajorTimescale**|
|**pjNoncritical**|**pjGanttMinorTimescale**|
|**pjCritical**|**pjBarTextLeft**|
|**pjMilestone**|**pjBarTextRight**|
|**pjSummary**|**pjBarTextTop**|
|**pjProjectSummary**|**pjBarTextBottom**|
|**pjMarked**|**pjBarTextInside**|
|**pjTaskFilterHighlight**|**pjGanttExternalTask**|
|**pjTaskRowColumnTitles**||

<br/>

#### If the Task Usage view is active

|||
|:-----|:-----|
|**pjAll**|**pjTaskFilterHighlight**|
|**pjCritical**|**pjTaskMajorTimescale**|
|**pjMarked**|**pjTaskMinorTimescale**|
|**pjMilestone**|**pjTaskRowColumnTitles**|
|**pjNoncritical**|**pjTaskUsageAssignmentRow**|
|**pjProjectSummary**|**pjTaskUsageExternalTask**|
|**pjSummary**||

<br/>

#### If the Task Sheet is active

|||
|:-----|:-----|
|**pjAll**|**pjGanttMajorTimescale**|
|**pjNoncritical**|**pjGanttMinorTimescale**|
|**pjCritical**|**pjBarTextLeft**|
|**pjMilestone**|**pjBarTextRight**|
|**pjSummary**|**pjBarTextTop**|
|**pjProjectSummary**|**pjBarTextBottom**|
|**pjMarked**|**pjBarTextInside**|
|**pjTaskFilterHighlight**|**pjGanttExternalTask**|
|**pjTaskRowColumnTitles**||

|||
|:-----|:-----|
|**pjAll**|**pjTaskFilterHighlight**|
|**pjCritical**|**pjTaskMajorTimescale**|
|**pjMarked**|**pjTaskMinorTimescale**|
|**pjMilestone**|**pjTaskRowColumnTitles**|
|**pjNoncritical**|**pjTaskUsageAssignmentRow**|
|**pjProjectSummary**|**pjTaskUsageExternalTask**|
|**pjSummary**||

|||
|:-----|:-----|
|**pjAll**|**pjProjectSummary**|
|**pjCritical**|**pjSummary**|
|**pjMarked**|**pjTaskSheetExternalTask**|
|**pjMilestone**|**pjTaskFilterHighlight**|
|**pjNoncritical**|**pjTaskRowColumnTitles**|

<br/>

#### PjColor constants

|||
|:-----|:-----|
|**pjColorAutomatic**|**pjNavy**|
|**pjAqua**|**pjOlive**|
|**pjBlack**|**pjPurple**|
|**pjBlue**|**pjRed**|
|**pjFuchsia**|**pjSilver**|
|**pjGray**|**pjTeal**|
|**pjGreen**|**pjYellow**|
|**pjLime**|**pjWhite**|
|**pjMaroon**||

<br/>

### Return value

 **Boolean**

## Remarks

Using the **TextStylesEx** method without specifying any arguments displays the **Text Styles** dialog box.

To set the text style by using hexadicimal RGB values, see the **[TextStyles32Ex](application-textstyles32ex-method-project.md)** method.
