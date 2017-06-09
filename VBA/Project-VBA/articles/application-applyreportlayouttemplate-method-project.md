---
title: Application.ApplyReportLayoutTemplate Method (Project)
keywords: vbapj.chm1524
f1_keywords:
- vbapj.chm1524
ms.prod: project-server
ms.assetid: cbc233c9-b955-3cd2-b1b8-99e4257bfea0
ms.date: 06/08/2017
---


# Application.ApplyReportLayoutTemplate Method (Project)
Applies the specified report template to the active report.

## Syntax

 _expression_. **ApplyReportLayoutTemplate** _(TemplateId)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TemplateId_|Optional|**[PjReportLayoutTemplateId](pjreportlayouttemplateid-enumeration-project.md)**|Specifies the kind of report; can be one of the following constants:  **pjReportLayoutComparison**,  **pjReportLayoutTitleAndChart**,  **pjReportLayoutTitleAndTable**, or  **pjReportLayoutTitleOnly**.|
| _TemplateId_|Optional|PJREPORTLAYOUTTEMPLATEID||

### Return value

 **Boolean**


## Remarks

For an existing report, the  **ApplyReportLayoutTemplate** method adds the specified report elements on top of other shapes in the report. For example, if the built-in Task Cost Overview report is active, the `ApplyReportLayoutTemplate pjReportLayoutTitleAndChart` statement adds a new text box with the report title and a new default chart to the report.


## Example

The following example creates a report that contains a title text box and a basic table, and then vertically centers text in the table cells.


```vb
Sub CreateTableReport()
    Dim theReport As Report
    Dim reportName As String
    Dim shp As Shape
    
    ' Add a report.
    reportName = "Table Report"
    Set theReport = ActiveProject.Reports.Add(reportName)
    
    ApplyReportLayoutTemplate TemplateId:=pjReportLayoutTitleAndTable
    
    For Each shp In theReport.Shapes
        If shp.HasTable Then
            shp.Select
            AlignTableCellVerticalCenter
        End If
    Next shp
End Sub
```


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[Report Object](report-object-project.md)
[PjReportLayoutTemplateId Enumeration](pjreportlayouttemplateid-enumeration-project.md)
