---
title: ViewsCombination.Add Method (Project)
keywords: vbapj.chm132807
f1_keywords:
- vbapj.chm132807
ms.prod: project-server
api_name:
- Project.ViewsCombination.Add
ms.assetid: 84e93698-88c3-b4a7-a754-8078fcab897a
ms.date: 06/08/2017
---


# ViewsCombination.Add Method (Project)

Adds a  **ViewCombination** object to a **ViewsCombination** collection.


## Syntax

 _expression_. **Add**( ** _Name_**, ** _TopView_**, ** _BottomView_**, ** _ShowInMenu_** )

 _expression_ A variable that represents a **ViewsCombination** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the combination view.|
| _TopView_|Required|**Variant**|The view that appears in the top pane of a combination view.|
| _BottomView_|Required|**Variant**|The view that appears in the bottom pane of a combination view.|
| _ShowInMenu_|Optional|**Boolean**|**True** if Project Server shows the view in the **View** menu. The default value is **False**|

### Return Value

 **ViewCombination**


## See also


#### Concepts


[ViewsCombination Collection Object](viewscombination-object-project.md)
