---
title: Report.Delete Method (Project)
keywords: vbapj.chm132549
f1_keywords:
- vbapj.chm132549
ms.prod: project-server
ms.assetid: 8a6b35c1-8552-b1be-2823-913790825a82
ms.date: 06/08/2017
---


# Report.Delete Method (Project)
Deletes the report.

## Syntax

 _expression_. **Delete**

 _expression_ A variable that represents a **Report** object.


### Return value

 **Nothing**


## Example

The following example determines whether a report named  **Report 1** exists, and if so, deletes the report. If the report is active, change to another view before you delete it; otherwise, Project shows run-time error 1004: **The table "Report 1" is in use and cannot be copied or deleted.**


```vb
Sub DeleteAReport()
    Dim reportName As String
    
    reportName = "Report 1"
    
    If ActiveProject.Reports.IsPresent(reportName) Then
        ' To delete the active report, change to another view.
        ViewApplyEx Name:="&;Gantt Chart"
        
        ActiveProject.Reports(reportName).Delete
    Else
        MsgBox Prompt:="No report name: " &; reportName, Title:="Report delete error"
    End If
End Sub
```


## See also


#### Other resources


[Report Object](report-object-project.md)
