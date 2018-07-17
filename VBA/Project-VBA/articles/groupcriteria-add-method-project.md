---
title: GroupCriteria.Add Method (Project)
ms.prod: project-server
api_name:
- Project.GroupCriteria.Add
ms.assetid: 6356acb9-0dbf-6e5e-e353-9673c3ed8097
ms.date: 06/08/2017
---


# GroupCriteria.Add Method (Project)

Adds a  **GroupCriterion** object to a **GroupCriteria** collection.


## Syntax

 _expression_. **Add**( ** _FieldName_**, ** _Ascending_**, ** _FontName_**, ** _FontSize_**, ** _FontBold_**, ** _FontItalic_**, ** _FontUnderLine_**, ** _FontColor_**, ** _CellColor_**, ** _Pattern_**, ** _GroupOn_**, ** _StartAt_**, ** _GroupInterval_** )

 _expression_ A variable that represents a **GroupCriteria** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldName_|Required|**String**|The name of the field being grouped by.|
| _Ascending_|Optional|**Boolean**|**True** if the field in a group definition should be grouped in ascending order. **False** if the field should be grouped in descending order. The default value is **True**.|
| _FontName_|Optional|**String**|The name of the font for a group definition.|
| _FontSize_|Optional|**Integer**|The size of the font in a group definition, in points. The default value is 8.|
| _FontBold_|Optional|**Boolean**|**True** if the font in a group definition is bold. The default value is **True**.|
| _FontItalic_|Optional|**Boolean**|**True** if the font in a group definition is italic. The default value is **False**.|
| _FontUnderLine_|Optional|**Boolean**|**True** if the font in a group definition is underlined. The default value is **False**.|
| _FontColor_|Optional|**Long**| The color of the font in a group definition. Can be one of the **[PjColor](pjcolor-enumeration-project.md)** constants. The default value is **pjBlack**.|
| _CellColor_|Optional|**Long**|The color of the cell background in a group definition. Can be one of the  **PjColor** constants. The default value is **pjColorAutomatic**.|
| _Pattern_|Optional|**Long**|The pattern for the cells in a group definition. Can be one of the  **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|
| _GroupOn_|Optional|**Long**|The type of grouping in a group definition. Can be one of the  **[PjGroupOn](pjgroupon-enumeration-project.md)** constants. The default value is **pjGroupOnEachValue**.|
| _StartAt_|Optional|**Variant**|The start of the intervals in a group definition. The default value is 0 for all fields except date fields, where it is the string "Project Start Date".|
| _GroupInterval_|Optional|**Variant**|The interval in a group definition. The default value is 1.|

### Return Value

 **GroupCriterion**


## See also


#### Concepts


[GroupCriteria Collection Object](groupcriteria-object-project.md)
