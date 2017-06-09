---
title: Report.Shapes Property (Project)
ms.prod: project-server
ms.assetid: 2f62c406-3845-79f8-3d17-e5891c1e23f9
ms.date: 06/08/2017
---


# Report.Shapes Property (Project)
Gets the collection of  **Shape** objects in the report. Read-only **Shapes**.

## Syntax

 _expression_. **Shapes**

 _expression_ A variable that represents a **Report** object.


## Example

The following example lists the shapes in a custom report. The report must be the active view to get the  **Shapes** collection; otherwise, you get a run-time error 424 (Object required) in the `For Each oShape In oReport.Shapes` statement.


```vb
Sub ListShapesInReport()
    Dim oReports As Reports
    Dim oReport As Report
    Dim oShape As shape
    Dim reportName As String
    Dim msg As String
    Dim msgBoxTitle As String
    Dim numShapes As Integer
    
    numShapes = 0
    msg = ""
    reportName = "New Table Tests"
    Set oReports = ActiveProject.Reports
    
    If oReports.IsPresent(reportName) Then
        ' Make the report the active view.
        oReports(reportName).Apply
        
        Set oReport = oReports(reportName)
        msgBoxTitle = "Shapes in report: '" &; oReport.Name &; "'"
    
        For Each oShape In oReport.Shapes
            numShapes = numShapes + 1
            msg = msg &; numShapes &; ". Shape type: " &; CStr(oShape.Type) _
                &; ", '" &; oShape.Name &; "'" &; vbCrLf
        Next oShape
        
        If numShapes > 0 Then
            MsgBox Prompt:=msg, Title:=msgBoxTitle
        Else
            MsgBox Prompt:="This report contains no shapes.", _
                Title:=msgBoxTitle
        End If
    Else
         MsgBox Prompt:="The requested report, '" &; reportName _
            &; "', does not exist.", Title:="Report error"
    End If
End Sub
```


## Property value

 **SHAPES**


## See also


#### Other resources


[Report Object](report-object-project.md)
[Shapes Object](shape-object-project.md)
