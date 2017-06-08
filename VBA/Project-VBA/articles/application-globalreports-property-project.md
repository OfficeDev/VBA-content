---
title: Application.GlobalReports Property (Project)
ms.prod: project-server
ms.assetid: 736be78c-2571-b07f-369c-845a06f9d1f9
ms.date: 06/08/2017
---


# Application.GlobalReports Property (Project)
Gets the collection of global (built-in) reports. Read-only  **Reports**.

## Syntax

 _expression_. **GlobalReports**

 _expression_ A variable that represents an **Application** object.


## Example

The following example prints a list of built-in reports in the  **Immediate** window of the VBE.


```vb
Sub ListGlobalReports()
    Dim oReport As Report

    Debug.Print "Number of global reports: " &; GlobalReports.Count
    
    For Each oReport In GlobalReports
        Debug.Print oReport.Index &; ": " &; oReport.Name
    Next oReport
End Sub
```

Following is the output for the RTM release of Project:




```
Number of global reports: 21
1: Project Overview
2: Burndown
3: Cost Overview
4: Work Overview
5: Task Cost Overview
6: Overallocated Resources
7: Upcoming Tasks
8: Earned Value Report
9: Cash Flow
10: Resource Cost Overview
11: Cost Overruns
12: Resource Overview
13: Milestone Report
14: Critical Tasks
15: Slipping Tasks
16: Late Tasks
17: Get started with Project
18: Create reports
19: Organize tasks
20: Share with your team
21: Best Practice Analyzer

```


## Property value

 **REPORTS**


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[Reports Object](reports-object-project.md)
