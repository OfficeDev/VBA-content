---
title: Application.FilterApply Method (Project)
keywords: vbapj.chm502
f1_keywords:
- vbapj.chm502
ms.prod: project-server
api_name:
- Project.Application.FilterApply
ms.assetid: d270862e-0577-a9db-e63b-9dcf1dc68b4a
ms.date: 06/08/2017
---


# Application.FilterApply Method (Project)

Sets the current filter.


## Syntax

 _expression_. **FilterApply**( ** _Name_**, ** _Highlight_**, ** _Value1_**, ** _Value2_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the filter to use.|
| _Highlight_|Optional|**Boolean**|**True** if Project highlights rows rather than applying the filter. The default value is **False**.|
| _Value1_|Optional|**String**|The first value to use when applying an interactive filter.|
| _Value2_|Optional|**String**|The second value to use when applying an interactive filter.|

### Return Value

 **Boolean**


## Example

The following example highlights filtered items.


```vb
Sub HighlightCriticalTasks() 
    FilterApply Name:="Critical", Highlight:=True 
End Sub
```


