---
title: Application.DetailsPaneToggle Method (Project)
keywords: vbapj.chm85
f1_keywords:
- vbapj.chm85
ms.prod: project-server
api_name:
- Project.Application.DetailsPaneToggle
ms.assetid: f62a42b2-397f-45c0-f2c1-f0468b8d489b
ms.date: 06/08/2017
---


# Application.DetailsPaneToggle Method (Project)

Shows or hides the task or resource  **Details** pane for the selected item.


## Syntax

 _expression_. **DetailsPaneToggle**( ** _Timeline_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Timeline_|Optional|**Variant**|If the value is  **True** or 1, the method shows or hides the project Timeline pane. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

If more than one task or resource is selected, the  **Details** pane shows the first item. If an empty item is selected, the project displays the Details pane with no data.


## Example

If the current view is a Calendar view where one task is selected, the following example first shows the  **Details** pane for that task, and then hides the **Details** pane and shows the **Timeline** pane.


```
DetailsPaneToggleDetailsPaneToggle Timeline:=True
```


