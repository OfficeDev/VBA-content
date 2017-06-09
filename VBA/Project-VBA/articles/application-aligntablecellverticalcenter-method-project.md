---
title: Application.AlignTableCellVerticalCenter Method (Project)
keywords: vbapj.chm1522
f1_keywords:
- vbapj.chm1522
ms.prod: project-server
ms.assetid: c790d8f7-e792-0718-3166-312640ff3f73
ms.date: 06/08/2017
---


# Application.AlignTableCellVerticalCenter Method (Project)
Aligns text at the vertical center of the cell, for selected cells in a report table.

## Syntax

 _expression_. **AlignTableCellVerticalCenter**

 _expression_ A variable that represents an **Application** object.


### Return value

 **Boolean**


## Example

In the following example, the  **AlignTableCells** macro aligns the text for all tables in the specified report.


```vb
Sub TestAlignReportTables()
    Dim reportName As String
    Dim alignment As String   ' The value can be "top", "center", or "bottom".
    
    reportName = "Align Table Cells Report"
    alignment = "top"
    
    AlignTableCells reportName, alignment
End Sub

' Align the text for all tables in a specified report.
Sub AlignTableCells(reportName As String, alignment As String)
    Dim theReport As Report
    Dim shp As Shape
    
    Set theReport = ActiveProject.Reports(reportName)
    
    ' Activate the report. If the report is already active,
    ' ignore the run-time error 1004 from the Apply method.
    On Error Resume Next
    theReport.Apply
    On Error GoTo 0
    
    For Each shp In theReport.Shapes
        Debug.Print "Shape: " &; shp.Type &; ", " &; shp.Name
        
        If shp.HasTable Then
            shp.Select
            
            Select Case alignment
                Case "top"
                    AlignTableCellTop
                Case "center"
                    AlignTableCellVerticalCenter
                Case "bottom"
                    AlignTableCellBottom
                Case Else
                    Debug.Print "AlignTableCells error: " &; vbCrLf _
                        &; "alignment must be top, center, or bottom."
                End Select
        End If
    Next shp
End Sub
```


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[Report Object](report-object-project.md)
[AlignTableCellTop Method](application-aligntablecelltop-method-project.md)
[AligntableCellBottom Method](application-aligntablecellbottom-method-project.md)
