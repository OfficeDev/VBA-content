---
title: Chart.Delete Method (Project)
keywords: vbapj.chm131614
f1_keywords:
- vbapj.chm131614
ms.prod: project-server
ms.assetid: 46312c6b-db7b-7562-d97a-d1fc8ba2acb7
ms.date: 06/08/2017
---


# Chart.Delete Method (Project)
Deletes a chart on an active report.

## Syntax

 _expression_. **Delete**

 _expression_ A variable that represents a **Chart** object.


### Return value

 **Variant**


## Example

The following example displays a report, and then deletes all charts on the report.


```vb
Sub DeleteCharts()
    Dim chartReport As Report
    Dim chartShape As Shape
    Dim reportName As String
    
    ' Display a report.
    reportName = "Chart Report 1"
    Set chartReport = ActiveProject.Reports(reportName)
    chartReport.Apply
    
    ' Delete every chart on the report.
    For Each chartShape In chartReport.Shapes
        If chartShape.Type = msoChart Then
            Debug.Print "Deleting chart: '" &; chartShape.Name _
                &; "' from report: " &; reportName
            chartShape.Delete
        End If
    Next chartShape
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
[Report.Delete Method](report-delete-method-project.md)
