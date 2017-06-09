---
title: Application.VisualReports Method (Project)
keywords: vbapj.chm2137
f1_keywords:
- vbapj.chm2137
ms.prod: project-server
api_name:
- Project.Application.VisualReports
ms.assetid: 4934cdcf-06b0-020c-3741-4ef70944cf98
ms.date: 06/08/2017
---


# Application.VisualReports Method (Project)

Opens the  **Visual Reports - Create Report** dialog box to the specified tab.


## Syntax

 _expression_. **VisualReports**( ** _PjVisualReportsTab_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PjVisualReportsTab_|Optional|**Long**|Specifies which tab to display. Can be one of the  **[PjVisualReportsTab](pjvisualreportstab-enumeration-project.md)** constants. The default is **pjTabAll**.|

### Return Value

 **Boolean**


## Remarks

The  **VisualReports** method returns **False** if successful.

The  **VisualReports** method corresponds to the **Visual Reports** command on the **REPORT** tab of the ribbon, which accesses the reports that use Excel and Visio templates. For the newer Office Art types of reports, see the **[ReportsDialog](application-reportsdialog-method-project.md)** method.


 **Note**  The  **[Reports](application-reports-method-project.md)** method, for the older style of reports that require connection with a printer, is deprecated in Project.


