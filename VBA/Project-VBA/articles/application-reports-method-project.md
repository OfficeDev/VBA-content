---
title: Application.Reports Method (Project)
keywords: vbapj.chm2334
f1_keywords:
- vbapj.chm2334
ms.prod: project-server
api_name:
- Project.Application.Reports
ms.assetid: 5288cc2d-538f-59c8-6c69-2244b1179cc1
ms.date: 06/08/2017
---


# Application.Reports Method (Project)

The  **Reports** method is deprecated in Project.


## Syntax

 _expression_. **Reports**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

The older style of reports that require connection with a printer are deprecated in Project. Running the  **Reports** method returns Run-time error 1100, "Application-defined or object-defined error".

For newer types of reports, see the  **[ReportsDialog](application-reportsdialog-method-project.md)** method for the Office Art types of reports or the **[VisualReports](application-visualreports-method-project.md)** method for the reports that use Excel and Visio templates.


