---
title: Application.ChangeStatusDate Method (Project)
keywords: vbapj.chm2181
f1_keywords:
- vbapj.chm2181
ms.prod: project-server
api_name:
- Project.Application.ChangeStatusDate
ms.assetid: 93635ef2-43c2-7cfd-5869-f8270a95a0ea
ms.date: 06/08/2017
---


# Application.ChangeStatusDate Method (Project)

Changes the project status date.


## Syntax

 _expression_. **ChangeStatusDate**( ** _Date_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Date_|Optional|**Variant**|New date for the project status date.|

### Return Value

 **Boolean**


## Remarks

The project status date enables Project to show progress lines in tasks on the Gantt chart. The status date is also used in earned value calculations. Using  **ChangeStatusDate** with no parameter shows the **Status Date** dialog box. If the user cancels the dialog box, **ChangeStatusDate** returns **False**.


## Example

The following example changes the project status date to April 7, 2012.


```
ChangeStatusDate("4/7/12")
```


