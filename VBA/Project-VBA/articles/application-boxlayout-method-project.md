---
title: Application.BoxLayout Method (Project)
keywords: vbapj.chm43
f1_keywords:
- vbapj.chm43
ms.prod: project-server
api_name:
- Project.Application.BoxLayout
ms.assetid: 4f26f5d1-41f2-56dc-e376-bcedd29613f9
ms.date: 06/08/2017
---


# Application.BoxLayout Method (Project)

Specifies the layout of boxes in the active Network Diagram view (PERT chart).


## Syntax

 _expression_. **BoxLayout**( ** _LayoutMode_**, ** _LayoutScheme_**, ** _SummaryPrecedence_**, ** _RowAlignment_**, ** _ColumnAlignment_**, ** _RowSpacing_**, ** _ColumnSpacing_**, ** _RowHeight_**, ** _ColumnWidth_**, ** _AdjustForPageBreaks_**, ** _ShowSummaryTasks_**, ** _ViewBackgroundColor_**, ** _ViewBackgroundPattern_**, ** _ShowProgressMarks_**, ** _ShowPageBreaks_**, ** _ShowIDOnly_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LayoutMode_|Optional|**Long**|Specifies whether the layout of boxes is controlled automatically or by the user, either with the  **LayoutNow** method or through the interface. Can be one of the **[PjLayoutMode](pjlayoutmode-enumeration-project.md)** constants.|
| _LayoutScheme_|Optional|**Long**|Specifies box alignment within each row. Can be one of the  **[PjLayoutScheme](pjlayoutscheme-enumeration-project.md)** constants.|
| _SummaryPrecedence_|Optional|**Boolean**|If  **True**, summary tasks are placed before subtasks.|
| _RowAlignment_|Optional|**Long**|Alignment of text within a row. Can be one of the  **[PjVerticalAlignment](pjverticalalignment-enumeration-project.md)** constants.|
| _ColumnAlignment_|Optional|**Long**|Alignment of text within a column. Can be one of the  **[PjAlignment](pjalignment-enumeration-project.md)** constants.|
| _RowSpacing_|Optional|**Long**|Spacing between rows. The value can be from 0 to 200.|
| _ColumnSpacing_|Optional|**Long**| Spacing between columns. The value can be from 0 to 200.|
| _RowHeight_|Optional|**Long**|The height of each row of boxes. Can be one of the  **[PjRowColSize](pjrowcolsize-enumeration-project.md)** constants.|
| _ColumnWidth_|Optional|**Long**|The width of each column of boxes. Can be one of the  **[PjRowColSize](pjrowcolsize-enumeration-project.md)** constants.|
| _AdjustForPageBreaks_|Optional|**Boolean**|If  **True**, a new task is placed on the next page if it does not fit on the current page. If **False**, a new task can fall on a break between pages.|
| _ShowSummaryTasks_|Optional|**Boolean**|If  **True**, summary tasks are shown. If **False**, summary tasks are hidden.|
| _ViewBackgroundColor_|Optional|**Long**|The background color of the view. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _ViewBackgroundPattern_|Optional|**Long**|The pattern used for the background. Can be one of the  **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|
| _ShowProgressMarks_|Optional|**Boolean**|**True** if tasks in progress are marked with a diagonal line from the upper-left corner of the box to the lower-right corner and completed tasks are marked with an additional diagonal line from the upper-right corner of the box to the lower-left corner. **False** if the progress of tasks is not marked.|
| _ShowPageBreaks_|Optional|**Boolean**|**True** if page breaks show in the Network Diagram; otherwise, **False**.|
| _ShowIDOnly_|Optional|**Boolean**|**True** if only task ID numbers are displayed. **False** if all the task data fields in Network Diagram boxes are displayed.|

### Return Value

 **Boolean**


## Remarks

Using the  **BoxLayout** method without specifying any arguments displays the **Box Layout** dialog box.

To format the Network Diagram layout using hexadecimal values for  _ViewBackgroundColor_, see the **[BoxLayoutEx](application-boxlayoutex-method-project.md)** method.


## Example

The following example sets the layout of boxes on the active Network Diagram view to the default values.


```vb
Sub ReturnToDefault() 
 Application.BoxLayout LayoutMode:=pjLayoutManual, LayoutScheme:=pjLayoutTopDownFromLeft, _ 
 SummaryPrecedence:=True, RowAlignment:=pjCenter, ColumnAlignment:=pjMiddle, RowSpacing:=45, _ 
 ColumnSpacing:=60, RowHeight:=pjSizeBestFit, ColumnWidth:=pjSizeBestFit, AdjustForPageBreaks:=True, _ 
 ShowSummaryTasks:=True, ViewBackgroundColor:=pjWhite, ViewBackgroundPattern:=pjBackgroundSolidFill, _ 
 ShowProgressMarks:=False, ShowPageBreaks:=True, ShowIDOnly:=False 
End Sub
```


