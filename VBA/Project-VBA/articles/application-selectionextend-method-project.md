---
title: Application.SelectionExtend Method (Project)
keywords: vbapj.chm2051
f1_keywords:
- vbapj.chm2051
ms.prod: project-server
api_name:
- Project.Application.SelectionExtend
ms.assetid: cffc56a0-0b25-2afa-427c-840aa2053921
ms.date: 06/08/2017
---


# Application.SelectionExtend Method (Project)

Turns selection extension on or off.


## Syntax

 _expression_. **SelectionExtend**( ** _Extend_**, ** _Add_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Extend_|Optional|**Boolean**|**True** if extend mode is active. (If extend mode is active, all items between the selection and the active item become part of the selection.) If **Extend** is **True**, **Add** is ignored. The default value is **False**.|
| _Add_|Optional|**Boolean**|**True** if add mode is active. (If add mode is active, only the active item is added to the selection.) The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example adds active item to the selection.


```vb
Sub Selection_Extend() 
 
 ViewApply Name:="&;Gantt Chart" 
 SelectionExtend Extend:=False, Add:=True 
 End Sub
```


