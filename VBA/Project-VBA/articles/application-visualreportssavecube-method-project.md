---
title: Application.VisualReportsSaveCube Method (Project)
keywords: vbapj.chm2139
f1_keywords:
- vbapj.chm2139
ms.prod: project-server
api_name:
- Project.Application.VisualReportsSaveCube
ms.assetid: 51b65e15-7ab5-79ff-9513-c47b204c1751
ms.date: 06/08/2017
---


# Application.VisualReportsSaveCube Method (Project)

Saves a Visual Reports cube to the default directory or to a specified directory.


## Syntax

 _expression_. **VisualReportsSaveCube**( ** _strNamePath_**, ** _PjVisualReportsCubeType_**, ** _ReportAlLFields_**, ** _PjVisualReportsDataLevel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _strNamePath_|Optional|**String**|Name and full path of the location to which to save the cube file (.cub).|
| _PjVisualReportsCubeType_|Optional|**Long**|Save cube type. Can be one of the  **[PjVisualReportsCubeType](pjvisualreportscubetype-enumeration-project.md)** consants. Default is **pjTaskTP**.|
| _ReportAlLFields_|Optional|**Boolean**|If  **True**, all noncustom fields are included in the report. Default is **False**.|
| _PjVisualReportsDataLevel_|Optional|**Long**|Save data level. Can be one of the  **[PjVisualReportsDataLevel](pjvisualreportsdatalevel-enumeration-project.md)** constants. Default is **pjLevelAutomatic**.|

### Return Value

 **Boolean**


## Remarks

The PjVisualReportsDataLevel parameter specifies the level to which the timephased data can be accessed. For example, if  **pjLevelMonths** (months) is specified, it not possible to access **pjLevelDays** (days).

Setting the ReportAllFields parameter to  **True** can degrade performance.


## Example

The following code saves a cube.


```vb
Sub a() 
 Dim tf As Boolean 
 tf = Application.VisualReportsSaveCube("c:\cube.cub", pjTaskNTP, , pjLevelQuarters) 
 If tf = True Then 
 MsgBox ("Cube saved successfully") 
 Else 
 MsgBox ("Cube not saved successfully") 
 End If 
End Sub
```


