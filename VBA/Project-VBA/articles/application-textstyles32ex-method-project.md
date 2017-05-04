---
title: Application.TextStyles32Ex Method (Project)
keywords: vbapj.chm2150
f1_keywords:
- vbapj.chm2150
ms.prod: PROJECTSERVER
api_name:
- Project.Application.TextStyles32Ex
ms.assetid: 8e1ed2bb-dac4-42d7-616b-a67984dcffa4
---


# Application.TextStyles32Ex Method (Project)

Sets the text styles for tasks and resources in the active view, where colors can be hexadecimal RGB values.


## Syntax

 _expression_. **TextStyles32Ex**( ** _Item_**, ** _Font_**, ** _Size_**, ** _Bold_**, ** _Italic_**, ** _Underline_**, ** _Color_**, ** _CellColor_**, ** _Pattern_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Integer**|The type of text to change. Can be one of the following  **[PjTextItem](pjtextitem-enumeration-project.md)** constants. If the **Gantt Chart** is active:
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
If the  **Task Usage** view is active:
|||
|:-----|:-----|
|**pjAll**|**pjTaskFilterHighlight**|
|**pjCritical**|**pjTaskMajorTimescale**|
|**pjMarked**|**pjTaskMinorTimescale**|
|**pjMilestone**|**pjTaskRowColumnTitles**|
|**pjNoncritical**|**pjTaskUsageAssignmentRow**|
|**pjProjectSummary**|**pjTaskUsageExternalTask**|
|**pjSummary**||
If the  **Task Sheet** is active:
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
|
| _Font_|Optional|**String**|The name of the font. The Font argument is ignored if the active view is the  **Network Diagram** and Item is not **pjAll**.|
| _Size_|Optional|**Integer**|The size of the font in points. The Size argument is ignored if the active view is the  **Network Diagram** and Item is not **pjAll**.|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the font. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|
| _CellColor_|Optional|**Long**|The background color of the cell. Can be a hexadecimal value for the RGB color.|
| _Pattern_|Optional|**Integer**|The background pattern of the cell. Can be one of the  **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **TextStyles32Ex** method without specifying any arguments displays the **Text Styles** dialog box.




 **Note**  If you use any of the  **PjColor** enumeration constants for the Color or CellColor parameters, the color will be nearly black. For example, the value of pjGreen is 9, which in the **TextStyles32Ex** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[TextStylesEx](application-textstylesex-method-project.md)** method.


