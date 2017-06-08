---
title: Application.ApplyReport Method (Project)
keywords: vbapj.chm2198
f1_keywords:
- vbapj.chm2198
ms.prod: project-server
ms.assetid: 869640a0-e45e-2e89-e3c9-ca15113ba8d3
ms.date: 06/08/2017
---


# Application.ApplyReport Method (Project)
Displays the specified report.

## Syntax

 _expression_. **ApplyReport** _(Name,_ _ApplyTo)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the report.|
| _ApplyTo_|Optional|**Variant**|The  _ApplyTo_ parameter is not used in Project.|
| _Name_|Optional|VARIANT||
| _ApplyTo_|Optional|VARIANT||

### Return value

 **Boolean**


## Remarks

The  **ApplyReport** method can be applied to custom reports and to built-in reports, such as Project Overview.


## Example

The following example checks whether a report exists; if so, the example displays the report, and then zooms the report to 80% of its original size.


```vb
Sub ReportZoom()
    Dim reportName As String
    reportName = "Report 1"
    
    If ActiveProject.Reports.IsPresent(reportName) Then
        ApplyReport reportName
        ZoomReport 80
    Else
        MsgBox Prompt:="No custom report name: " &; reportName, Title:="Report apply error"
    End If
End Sub
```


## See also


#### Other resources


[Report.Apply Method](report-apply-method-project.md)
