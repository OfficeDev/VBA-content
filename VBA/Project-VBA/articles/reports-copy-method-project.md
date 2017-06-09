---
title: Reports.Copy Method (Project)
ms.prod: project-server
ms.assetid: fd930e98-4200-05e0-67e3-f4d34ae26928
ms.date: 06/08/2017
---


# Reports.Copy Method (Project)
Copies a custom report and creates a new report with the same content.

## Syntax

 _expression_. **Copy** _(Source,_ _NewName)_

 _expression_ A variable that represents a **Reports** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**Variant**|Name or  **Report** object of the report to copy.|
| _NewName_|Required|**String**|Name of the new report.|
| _Source_|Required|VARIANT||
| _NewName_|Required|STRING||

### Return value

 **Report**

The new report.


## Example

The  **CopyAReport** macro checks whether the specified report to copy exists, and checks whether the new report already exists. The macro then uses one of the variants of the _Source_ parameter to create a copy of the report, and then displays the new report.


```vb
Sub CopyAReport()
    Dim reportName As String
    Dim newReportName As String
    Dim newExists As Boolean
    Dim oldExists As Boolean
    Dim report2Copy As Report
    Dim newReport As Report
    
    reportName = "Table Tests"
    newReportName = "New Table Tests"
    oldExists = ActiveProject.Reports.IsPresent(reportName)
    newExists = ActiveProject.Reports.IsPresent(newReportName)
    
    Debug.Print "oldExists " &; CStr(oldExists) &; "; newExists " &; newExists
    
    If oldExists And Not newExists Then
        Set report2Copy = ActiveProject.Reports(reportName)
        
        ' You can use either of the following two statements.
        'Set newReport = ActiveProject.Reports.Copy(report2Copy, newReportName)
        Set newReport = ActiveProject.Reports.Copy(reportName, newReportName)
       
        newReport.Apply
    End If
    
    If (oldExists = False) Then
         MsgBox Prompt:="The requested report to copy, '" &; reportName _
            &; "', does not exist.", Title:="Report copy error"
    ElseIf newExists Then
        MsgBox Prompt:="The new report '" &; newReportName _
            &; "' already exists.", Title:="Report copy error"
    Else
        MsgBox Prompt:="The new report '" &; newReportName &; "'" _
            &; vbCrLf &; "is copied from '" &; reportName &; "'.", _
            Title:="Report copy success"
    End If
End Sub
```


## See also


#### Other resources


[Reports Object](reports-object-project.md)
[Report Object](report-object-project.md)
