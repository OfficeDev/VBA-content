---
title: MultiSelect Property (Microsoft Forms)
keywords: fm20.chm5225069
f1_keywords:
- fm20.chm5225069
ms.prod: office
ms.assetid: 4c8102d4-abbb-a7f7-8dd3-0a0695752fa8
ms.date: 06/08/2017
---


# MultiSelect Property (Microsoft Forms)



Indicates whether the object permits multiple selections.
 **Syntax**
 _object_. **MultiSelect** [= _fmMultiSelect_ ]
The  **MultiSelect** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmMultiSelect_|Optional. The selection mode that the control uses.|
 **Settings**
The settings for  _fmMultiSelect_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmMultiSelectSingle_|0|Only one item can be selected (default).|
| _fmMultiSelectMulti_|1|Pressing the SPACEBAR or clicking selects or deselects an item in the list.|
| _fmMultiSelectExtended_|2|Pressing SHIFT and clicking the mouse, or pressing SHIFT and one of the arrow keys, extends the selection from the previously selected item to the current item. Pressing CTRL and clicking the mouse selects or deselects an item.|
 **Remarks**
When the  **MultiSelect** property is set to _Extended_ or _Simple_, you must use the list box's **Selected** property to determine the selected items. Also, the **Value** property of the control is always **Null**.
The  **ListIndex** property returns the index of the row with the keyboard[focus](vbe-glossary.md).

