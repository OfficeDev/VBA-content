---
title: GroupCriteria2.Add Method (Project)
ms.prod: project-server
api_name:
- Project.GroupCriteria2.Add
ms.assetid: c10914c1-eda2-128e-0c5d-056ee51a9076
ms.date: 06/08/2017
---


# GroupCriteria2.Add Method (Project)

Adds a  **GroupCriterion2** object to the **GroupCriteria2** collection.


## Syntax

 _expression_. **Add**( ** _FieldName_**, ** _Ascending_**, ** _FontName_**, ** _FontSize_**, ** _FontBold_**, ** _FontItalic_**, ** _FontUnderLine_**, ** _FontColor_**, ** _CellColor_**, ** _Pattern_**, ** _GroupOn_**, ** _StartAt_**, ** _GroupInterval_** )

 _expression_ An expression that returns a **GroupCriteria2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldName_|Required|**String**|The name of the field being grouped by.|
| _Ascending_|Optional|**Boolean**|**True** if the field in a group definition should be grouped in ascending order. **False** if the field should be grouped in descending order. The default value is **True**.|
| _FontName_|Optional|**String**|The name of the font for a group definition.|
| _FontSize_|Optional|**[INT]**|The size of the font in a group definition, in points. The default value is 8.|
| _FontBold_|Optional|**Boolean**|**True** if the font in a group definition is bold. The default value is **True**.|
| _FontItalic_|Optional|**Boolean**|**True** if the font in a group definition is italic. The default value is **False**.|
| _FontUnderLine_|Optional|**Boolean**|**True** if the font in a group definition is underlined. The default value is **False**.|
| _FontColor_|Optional|**PjColor**| The color of the font in a group definition. Can be one of the **[PjColor](pjcolor-enumeration-project.md)** constants. The default value is **pjBlack**.|
| _CellColor_|Optional|**PjColor**|The color of the cell background in a group definition. Can be one of the  **PjColor** constants. The default value is **pjColorAutomatic**.|
| _Pattern_|Optional|**PjBackgroundPattern**|The pattern for the cells in a group definition. Can be one of the  **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|
| _GroupOn_|Optional|**PjGroupOn**|The type of grouping in a group definition. Can be one of the  **[PjGroupOn](pjgroupon-enumeration-project.md)** constants. The default value is **pjGroupOnEachValue**.|
| _StartAt_|Optional|**Variant**|The start of the intervals in a group definition. The default value is 0 for all fields except date fields, where it is the string "Project Start Date".|
| _GroupInterval_|Optional|**Variant**|The interval in a group definition. The default value is 1.|

### Return Value

GroupCriterion2


## Remarks

To add a  **GroupCriterion2** object where colors can be hexadecimal values, use the **[AddEx](groupcriteria2-addex-method-project.md)** method.


## See also


#### Concepts


[GroupCriteria2 Collection Object](groupcriteria2-object-project.md)

