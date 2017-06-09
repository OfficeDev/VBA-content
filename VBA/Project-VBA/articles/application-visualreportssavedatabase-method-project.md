---
title: Application.VisualReportsSaveDatabase Method (Project)
keywords: vbapj.chm2138
f1_keywords:
- vbapj.chm2138
ms.prod: project-server
api_name:
- Project.Application.VisualReportsSaveDatabase
ms.assetid: edcbaff5-beb1-ba11-fb65-ec26a24ab23d
ms.date: 06/08/2017
---


# Application.VisualReportsSaveDatabase Method (Project)

Saves a Visual Reports database to the default directory or to a specified directory.


## Syntax

 _expression_. **VisualReportsSaveDatabase**( ** _strNamePath_**, ** _PjVisualReportsDataLevel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _strNamePath_|Optional|**String**|Name and full path of the location to which to save the database file (.mbd).|
| _PjVisualReportsDataLevel_|Optional|**Long**|Save data level. Can be one of the  **[PjVisualReportsDataLevel](pjvisualreportsdatalevel-enumeration-project.md)** constants. Default is **pjLevelAutomatic**.|

### Return Value

 **Boolean**


## Remarks

The PjVisualReportsDataLevel parameter specifies the level to which the timephased data can be accessed. For example, if  **pjLevelMonths** (months) is specified, it not possible to access **pjLevelDays** (days).


## Example

Following is an example of using The  **VisualReportsSaveDatabase** method.


```vb
Sub a() 
 Dim tf As Boolean 
 tf = Application.VisualReportsSaveDatabase("C:\mydb.mdb", pjLevelAutomatic) 
 If tf = True Then 
 MsgBox ("Database saved successfully") 
 Else 
 MsgBox ("Database wasn't saved successfully") 
 End If 
End Sub
```


