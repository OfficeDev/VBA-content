---
title: Reports Object (Project)
ms.prod: project-server
ms.assetid: a9f4a13b-1907-dbe8-8077-fb1226bb8bb9
ms.date: 06/08/2017
---


# Reports Object (Project)
Contains a collection of  **[Report](report-object-project.md)** objects, where each report is a custom report.
 

## Example

The  **Reports** object is the collection of custom reports in a project. It does not include the built-in reports, such as **Project Overview**,  **Critical Tasks**, or  **Milestone Report**. Use the  **Project.Reports** property to get the **Reports** collection object, as in the following example:
 

 

```
Sub ListCustomReports()
    Dim oReport As Report
    Dim msg As String
    Dim msgBoxTitle As String
    msg = ""
    msgBoxTitle = "Custom reports in '" &amp; ActiveProject.Name &amp; "'"
    
    For Each oReport In ActiveProject.Reports
        msg = msg &amp; oReport.Index &amp; ": " &amp; oReport.Name &amp; vbCrLf
    Next oReport
        
    If ActiveProject.Reports.Count > 0 Then
        MsgBox Prompt:=msg, Title:=msgBoxTitle
    Else
        MsgBox Prompt:="This project contains no custom reports.", _
            Title:=msgBoxTitle
    End If
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](reports-add-method-project.md)|
|[Copy](reports-copy-method-project.md)|
|[IsPresent](reports-ispresent-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](reports-application-property-project.md)|
|[Count](reports-count-property-project.md)|
|[Item](reports-item-property-project.md)|
|[Parent](reports-parent-property-project.md)|

## See also


#### Other resources


 
[Report Object](report-object-project.md)
 
[Project.Reports Property](project-reports-property-project.md)
