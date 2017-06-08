---
title: Selected Property
keywords: fm20.chm2001830
f1_keywords:
- fm20.chm2001830
ms.prod: office
api_name:
- Office.Selected
ms.assetid: 5a286e96-d250-089a-1682-da00112157aa
ms.date: 06/08/2017
---


# Selected Property



Returns or sets the selection state of items in a  **ListBox**.
 **Syntax**
 _object_. **Selected(**_index_**)** [= _Boolean_ ]
The  **Selected** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. An integer with a range from 0 to one less than the number of items in the list.|
| _Boolean_|Optional. Whether an item is selected.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The item is selected.|
|**False**|The item is not selected.|
 **Remarks**
The  **Selected** property is useful when users can make multiple selections. You can use this property to determine the selected rows in a multi-select list box. You can also use this property to select or deselect rows in a list from code.
The default value of this property is based on the current selection state of the  **ListBox**.
For single-selection list boxes, the  **Value** or **ListIndex** properties are recommended for getting and setting the selection. In this case, **ListIndex** returns the index of the selected item. However, in a multiple selection, **ListIndex** returns the index of the row contained within the[focus](vbe-glossary.md) rectangle, regardless of whether the row is actually selected.
When a list box control's  **MultiSelect** property is set to _None_, only one row can have its **Selected** property set to **True**.
Entering a value that is out of range for the index does not generate an error message, but does not set a property for any item in the list.

