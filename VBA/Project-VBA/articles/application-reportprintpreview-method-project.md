---
title: Application.ReportPrintPreview Method (Project)
keywords: vbapj.chm112
f1_keywords:
- vbapj.chm112
ms.prod: project-server
api_name:
- Project.Application.ReportPrintPreview
ms.assetid: f93003ee-c25e-9581-191e-478bb30314f0
ms.date: 06/08/2017
---


# Application.ReportPrintPreview Method (Project)

Deprecated in Project. Shows an on-screen preview of a printed report.


## Syntax

 _expression_. **ReportPrintPreview**( ** _Name_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a report for which to show an on-screen preview. |

### Return Value

 **Boolean**


## Remarks

In Project, the  **ReportPrintPreview** method returns error 1100, "The method is not available in this situation." In Project, if you execute the **ReportPrintPreview** method with no argument, it displays the **Custom Reports** dialog box.


