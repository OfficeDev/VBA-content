---
title: Report.Apply Method (Project)
ms.prod: project-server
ms.assetid: 4461da82-5bd6-2d9b-0d39-35875c2cee36
ms.date: 06/08/2017
---


# Report.Apply Method (Project)
Changes the view to display the report.

## Syntax

 _expression_. **Apply**

 _expression_ A variable that represents a **Report** object.


### Return value

 **Nothing**


## Example

The following example determines whether a report named  **Report 1** exists, and if so, displays the report.


```vb
Sub ShowAReport()
    Dim reportName As String
    
    reportName = "Report 1"
    
    If ActiveProject.Reports.IsPresent(reportName) Then
        ActiveProject.Reports(reportName).Apply
    Else
        MsgBox Prompt:="No report name: " &; reportName, Title:="Report apply error"
    End If
End Sub
```


## See also


#### Other resources


[Report Object](report-object-project.md)
