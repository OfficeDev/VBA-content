---
title: Application.VisualReportsEdit Method (Project)
keywords: vbapj.chm2143
f1_keywords:
- vbapj.chm2143
ms.prod: project-server
api_name:
- Project.Application.VisualReportsEdit
ms.assetid: ba439985-f18b-f9a3-23d5-3d5ae39c50dc
ms.date: 06/08/2017
---


# Application.VisualReportsEdit Method (Project)

Opens the default or a specified Visual Reports template for editing.


## Syntax

 _expression_. **VisualReportsEdit**( ** _strVisualReportTemplateFile_**, ** _PjVisualReportsDataLevel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _strVisualReportTemplateFile_|Optional|**String**|Full path and the name of template file.|
| _PjVisualReportsDataLevel_|Optional|**Long**|Data level for the template. Can be one of the  **[PjVisualReportsDataLevel](pjvisualreportsdatalevel-enumeration-project.md)** constants. The default is **pjLevelAutomatic**.|

### Return Value

 **Boolean**


## Remarks

The PjVisualReportsDataLevel parameter specifies the level to which the timephased data can be accessed. For example, if  **pjLevelMonths** (months) is specified, it not possible to access **pjLevelDays** (days).


## Example

The following example opens the "MyTemplate.xlt" template, with a data level of months.


```vb
Application.VisualReportsEdit("C:\MyTemplate.xlt", pjMonths)
```


