---
title: Application.ZoomReport Method (Project)
keywords: vbapj.chm2196
f1_keywords:
- vbapj.chm2196
ms.prod: project-server
ms.assetid: 05a0ec6e-1329-2545-df89-5d87af88a454
ms.date: 06/08/2017
---


# Application.ZoomReport Method (Project)
Zooms (enlarges or shrinks) the active report to the specified percentage of its original size.

## Syntax

 _expression_. **ZoomReport** _(Percent,_ _Entire)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Percent_|Optional|**Variant**|The percent of the original size.|
| _Entire_|Optional|**Variant**|The  _Entire_ parameter has no effect.|
| _Percent_|Optional|VARIANT||
| _Entire_|Optional|VARIANT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Boolean**


## Remarks

The  _Percent_ parameter can have a value of 10 to 400. If the value is outside of that range, the **ZoomReport** method shows a run-time error 1101, "The argument value is not valid."

The  **ZoomReport** method can be applied to custom reports and to built-in reports, such as Project Overview. When you change the report size, switch to another view, and then return to the previous report, the zoom level remains in effect. To restore the original size, use the following command: `ZoomReport 100`.


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


