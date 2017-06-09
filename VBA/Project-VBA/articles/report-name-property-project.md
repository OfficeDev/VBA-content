---
title: Report.Name Property (Project)
ms.prod: project-server
ms.assetid: da13696d-313a-3d78-2f1b-34d5fea4c2a9
ms.date: 06/08/2017
---


# Report.Name Property (Project)
Gets or sets the name of the report. Read/write  **String**.

## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Report** object.


## Example

The following example lists the index and name of each custom report in a project.


```vb
Sub ListCustomReports()
    Dim oReport As Report
    Dim msg As String
    Dim msgBoxTitle As String
    msg = ""
    msgBoxTitle = "Custom reports in '" &; ActiveProject.Name &; "'"
    
    For Each oReport In ActiveProject.Reports
        msg = msg &; oReport.Index &; oReport.Name &; vbCrLf
    Next oReport
        
    If ActiveProject.Reports.Count > 0 Then
        MsgBox Prompt:=msg, Title:=msgBoxTitle
    Else
        MsgBox Prompt:="This project contains no custom reports.", _
            Title:=msgBoxTitle
    End If
End Sub
```


## Property value

 **STRING**


## See also


#### Other resources


[Report Object](report-object-project.md)
[Reports Object](reports-object-project.md)
