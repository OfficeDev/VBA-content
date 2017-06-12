---
title: Filter.Apply Method (Project)
keywords: vbapj.chm132210
f1_keywords:
- vbapj.chm132210
ms.prod: project-server
api_name:
- Project.Filter.Apply
ms.assetid: bc9a406c-d4ae-0fa5-a5b1-70bf3520fac4
ms.date: 06/08/2017
---


# Filter.Apply Method (Project)

Applies the filter to the current view.


## Syntax

 _expression_. **Apply**( ** _Highlight_** )

 _expression_ An expression that returns a **Filter** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Highlight_|Optional|**Boolean**|If  **True**, highlights the filtered items within the list of all items. If **False**, shows only the filtered items in the view. The default is **False**.|

### Return Value

 **Nothing**


## Example

If the current view is a task view, the following example highlights the critical tasks. 


```vb
ActiveProject.TaskFilters("Critical").Apply Highlight:=True
```


