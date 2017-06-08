---
title: Application.Sort Method (Project)
keywords: vbapj.chm903
f1_keywords:
- vbapj.chm903
ms.prod: project-server
api_name:
- Project.Application.Sort
ms.assetid: 996df315-32ae-eac8-75cb-182a95f74879
ms.date: 06/08/2017
---


# Application.Sort Method (Project)

Sorts the tasks or resources in the active pane.


## Syntax

 _expression_. **Sort**( ** _Key1_**, ** _Ascending1_**, ** _Key2_**, ** _Ascending2_**, ** _Key3_**, ** _Ascending3_**, ** _Renumber_**, ** _Outline_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Key1_|Optional|**String**|The name of a primary field to sort. If Key1 is omitted, Project displays the  **Sort** dialog box.|
| _Ascending1_|Optional|**Boolean**|**True** if the primary field will be sorted in ascending order. The default value is **True**.|
| _Key2_|Optional|**String**|The name of a secondary field to sort.|
| _Ascending2_|Optional|**Boolean**|**True** if the secondary field will be sorted in ascending order. The default value is **True.**|
| _Key3_|Optional|**String**|The name of a tertiary field to sort.|
| _Ascending3_|Optional|**Boolean**|**True** if the tertiary field will be sorted in ascending order. The default value is **True**.|
| _Renumber_|Optional|**Boolean**|**True** if Project renumbers tasks after sorting them. For task views, Renumber can be **True** only if Outline is **True**. If Outline is **True**, then Renumber defaults to the current setting in the **Sort** dialog box. If Outline is **False**, Renumber is ignored.|
| _Outline_|Optional|**Boolean**|**True** if the outline level of tasks or resources will be preserved after sorting them. The default value is **True**.|

### Return Value

 **Boolean**


## Example

The following example sorts the tasks in the active project by priority, and then renumbers the tasks.


```vb
Sub SortByPriority() 
 Sort Key1:="Priority", Ascending1:=True, Renumber:=True 
End Sub
```


