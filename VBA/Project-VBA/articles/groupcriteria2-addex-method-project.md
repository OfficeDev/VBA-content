---
title: GroupCriteria2.AddEx Method (Project)
keywords: vbapj.chm132308
f1_keywords:
- vbapj.chm132308
ms.prod: project-server
api_name:
- Project.GroupCriteria2.AddEx
ms.assetid: 8474aa63-bf63-be29-86ef-177d8105e105
ms.date: 06/08/2017
---


# GroupCriteria2.AddEx Method (Project)

Adds a  **GroupCriterion2** object to the **GroupCriteria2** collection, where **CellColor** can be a hexadecimal value.


## Syntax

 _expression_. **AddEx**( ** _FieldName_**, ** _Ascending_**, ** _FontName_**, ** _FontSize_**, ** _FontBold_**, ** _FontItalic_**, ** _FontUnderLine_**, ** _FontColor_**, ** _CellColor_**, ** _Pattern_**, ** _GroupOn_**, ** _StartAt_**, ** _GroupInterval_** )

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
| _FontColor_|Optional|**Long**| The color of the font in a group definition, where color can be a hexadecimal value. See remarks. The default value is &;H0.|
| _CellColor_|Optional|**Long**|The color of the cell background specified by a hexadecimal value. See remarks. The default value is &;HFFFFFF.|
| _Pattern_|Optional|**PjBackgroundPattern**|The pattern for the cells in a group definition. Can be one of the  **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|
| _GroupOn_|Optional|**PjGroupOn**|The type of grouping in a group definition. Can be one of the  **[PjGroupOn](pjgroupon-enumeration-project.md)** constants. The default value is **pjGroupOnEachValue**.|
| _StartAt_|Optional|**Variant**|The start of the intervals in a group definition. The default value is 0 for all fields except date fields, where it is the string "Project Start Date".|
| _GroupInterval_|Optional|**Variant**|The interval in a group definition. The default value is 1.|

### Return Value

 **GroupCriterion2**


## Remarks

RGB colors can be expressed in decimal or hexadecimal values. In Project, red is the last byte of a hexadecimal value. For example, if the value of CellColorEx is 65535, the color is blue (&;HFF0000). 

The valid range for a normal RGB color is 0 to 16,777,215 (&;HFFFFFF&;). Each color setting (property or argument) is a 4-byte integer. The high byte of a number in this range equals 0. The lower 3 bytes, from least to most significant byte, determine the amount of red, green, and blue, respectively. The red, green, and blue components are each represented by a number between 0 and 255 (&;HFF). 


## See also


#### Concepts


[GroupCriteria2 Collection Object](groupcriteria2-object-project.md)

