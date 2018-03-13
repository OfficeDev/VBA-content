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



| <strong>Name</strong>  | <strong>Required/Optional</strong> | <strong>Data Type</strong> | <strong>Description</strong>                                                                                                                                                                                                                        |
|:-----------------------|:-----------------------------------|:---------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>Name</em>          | Required                           | <strong>String</strong>    | The name of the view.                                                                                                                                                                                                                               |
| <em>Screen</em>        | Optional                           | <strong>Long</strong>      | The project view. Can be one of the <strong><a href="pjviewscreen-enumeration-project.md" data-raw-source="[PjViewScreen](pjviewscreen-enumeration-project.md)">PjViewScreen</a></strong> constants. The default value is <strong>pjGantt</strong>. |
| <em>ShowInMenu</em>    | Optional                           | <strong>Boolean</strong>   | <strong>True</strong> if Project Server adds the single-pane view to the <strong>View</strong> menu. The default value is <strong>False</strong>.                                                                                                   |
| <em>Table</em>         | Optional                           | <strong>Variant</strong>   | Specifies the table to be used by the view. This value is ignored if the view specified with the  <strong>Screen</strong> argument does not use tables.                                                                                             |
| <em>Filter</em>        | Optional                           | <strong>Variant</strong>   | Specifies the filter to be used on the view.                                                                                                                                                                                                        |
| <em>Group</em>         | Optional                           | <strong>Variant</strong>   | Specifies the group to be used by the view. If a group is required for the view, but none is specified, the default is ** No Group<strong>. This value is ignored if the view specified with the **Screen</strong> argument does not use groups.    |
| <em>HighlightFilt</em> | Optional                           | <strong>Boolean</strong>   | <strong>True</strong> if the filter applied is a highlight filter. The default value is <strong>False</strong>.                                                                                                                                     |

### Return Value

 **ViewSingle**


## See also


#### Concepts


[ViewsSingle Collection Object](viewssingle-object-project.md)
