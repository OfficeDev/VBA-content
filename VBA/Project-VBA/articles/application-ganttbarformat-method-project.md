---
title: Application.GanttBarFormat Method (Project)
keywords: vbapj.chm938
f1_keywords:
- vbapj.chm938
ms.prod: project-server
api_name:
- Project.Application.GanttBarFormat
ms.assetid: 2b3b3933-1993-d4cf-f4ff-475c4b003514
ms.date: 06/08/2017
---


# Application.GanttBarFormat Method (Project)

Changes the formatting of Gantt bars from their default styles.


## Syntax

 _expression_. **GanttBarFormat**( ** _TaskID_**, ** _GanttStyle_**, ** _StartShape_**, ** _StartType_**, ** _StartColor_**, ** _MiddleShape_**, ** _MiddlePattern_**, ** _MiddleColor_**, ** _EndShape_**, ** _EndType_**, ** _EndColor_**, ** _LeftText_**, ** _RightText_**, ** _TopText_**, ** _BottomText_**, ** _InsideText_**, ** _Reset_**, ** _ProjectName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Optional|**Long**|The identification number of the task to be changed on the Gantt chart. The default is to change the Gantt bars of the selected tasks.|
| _GanttStyle_|Optional|**Integer**|The style applied to the Gantt bar to be formatted. The value for GanttStyle is based on the position of the bar style in the  **Bar Styles** dialog box. For example, the value 3 returns the third bar style in the list.|
| _StartShape_|Optional|**Long**|The start shape of the Gantt bar. Can be one of the  **[PjBarEndShape](pjbarendshape-enumeration-project.md)** constants.|
| _StartType_|Optional|**Long**|The start type of the Gantt bar. Can be one of the  **[PjBarType](pjbartype-enumeration-project.md)** constants.|
| _StartColor_|Optional|**Long**|The color of the start shape of the Gantt bar. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _MiddleShape_|Optional|**Long**|The middle shape of the Gantt bar. Can be one of the  **[PjBarShape](pjbarshape-enumeration-project.md)** constants.|
| _MiddlePattern_|Optional|**Long**|The middle pattern of the Gantt bar. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _MiddleColor_|Optional|**Long**|The color of the middle section Gantt bar. Can be one of the  **PjColor** constants.|
| _EndShape_|Optional|**Long**|The end shape of the Gantt bar. Can be one of the  **PjBarEndShape** constants.|
| _EndType_|Optional|**Long**|The end type of the Gantt bar. Can be one of the following  **PjBarType** constants: **pjDashed**, **pjFramed**, or **pjSolid**.|
| _EndColor_|Optional|**Long**|The color of the end shape of the Gantt bar. Can be one of the  **PjColor** constants.|
| _LeftText_|Optional|**String**|The task field to display to the left of the Gantt bar.|
| _RightText_|Optional|**String**|The task field to display to the right of the Gantt bar.|
| _TopText_|Optional|**String**|The task field to display above the Gantt bar.|
| _BottomText_|Optional|**String**|The task field to display below the Gantt bar.|
| _InsideText_|Optional|**String**|The task field to display inside the Gantt bar.|
| _Reset_|Optional|**Boolean**|**True** if the bar formatting is reset to the default formatting of the style in the **Bar Styles** dialog box; otherwise, **False**.|
| _ProjectName_|Optional|**String**|The name of the project containing  **TaskID** if consolidation is involved. The default value is the name of the active project.|

### Return Value

 **Boolean**


## Remarks

Using the  **GanttBarFormat** method without specifying any arguments displays the **Format Bar** dialog box.

 To define the default styles, use the **[GanttBarStyleEdit](application-ganttbarstyleedit-method-project.md)** method.

To change Gantt bar formatting where colors can be hexadecimal RGB values, use the  **[GanttBarFormatEx](application-ganttbarformatex-method-project.md)** method.


## Example

The following example displays a red diamond shape for the start of the task with the Task ID of 3.


```vb
Sub GanttBar_Format() 
 
 'Activate Gantt Chart view 
 ViewApply Name:="&;Gantt Chart" 
 GanttBarFormat TaskID:=3, StartShape:=pjDiamond, StartType:=pjSolid, StartColor:=pjRed 
 
End Sub
```


