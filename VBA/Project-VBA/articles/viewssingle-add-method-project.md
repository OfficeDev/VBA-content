---
title: ViewsSingle.Add Method (Project)
keywords: vbapj.chm132759
f1_keywords:
- vbapj.chm132759
ms.prod: project-server
api_name:
- Project.ViewsSingle.Add
ms.assetid: 509103f7-6301-0880-75eb-590141179caf
ms.date: 06/08/2017
---


# ViewsSingle.Add Method (Project)

Adds a  **ViewSingle** object to a **ViewsSingle** collection.


## Syntax

 _expression_. **Add**( ** _Name_**, ** _Screen_**, ** _ShowInMenu_**, ** _Table_**, ** _Filter_**, ** _Group_**, ** _HighlightFilt_** )

 _expression_ A variable that represents a **ViewsSingle** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the view.|
| _Screen_|Optional|**Long**| The project view. Can be one of the **[PjViewScreen](pjviewscreen-enumeration-project.md)** constants. The default value is **pjGantt**.|
| _ShowInMenu_|Optional|**Boolean**|**True** if Project Server adds the single-pane view to the **View** menu. The default value is **False**.|
| _Table_|Optional|**Variant**|Specifies the table to be used by the view. This value is ignored if the view specified with the  **Screen** argument does not use tables.|
| _Filter_|Optional|**Variant**|Specifies the filter to be used on the view.|
| _Group_|Optional|**Variant**|Specifies the group to be used by the view. If a group is required for the view, but none is specified, the default is ** No Group**. This value is ignored if the view specified with the **Screen** argument does not use groups.|
| _HighlightFilt_|Optional|**Boolean**|**True** if the filter applied is a highlight filter. The default value is **False**.|

### Return Value

 **ViewSingle**


## See also


#### Concepts


[ViewsSingle Collection Object](viewssingle-object-project.md)
