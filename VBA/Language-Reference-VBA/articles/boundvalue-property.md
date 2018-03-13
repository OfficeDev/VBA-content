---
title: BoundValue Property
keywords: fm20.chm5225012
f1_keywords:
- fm20.chm5225012
ms.prod: office
api_name:
- Office.BoundValue
ms.assetid: a064f85f-981c-f710-393c-05f14c00249d
ms.date: 06/08/2017
---


# BoundValue Property



Contains the value of a control when that control receives the focus.
 **Syntax**
 _object_. **BoundValue** [= _Variant_ ]
The  **BoundValue** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                           |
|:----------------------|:-------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                              |
| <em>Variant</em>      | Optional. The current state or content of the control. |

 **Settings**


| <strong>Control</strong>           | <strong>Description</strong>                                                                               |
|:-----------------------------------|:-----------------------------------------------------------------------------------------------------------|
| <strong>CheckBox</strong>          | An integer value indicating whether the item is selected:                                                  |
|                                    | Null Indicates the item is in a null state, neither selected nor [cleared](glossary-vba.md).               |
|                                    | -1 True. Indicates the item is selected.                                                                   |
|                                    | 0 False. Indicates the item is cleared.                                                                    |
| <strong>OptionButton</strong>      | Same as  <strong>CheckBox</strong>.                                                                        |
| <strong>ToggleButton</strong>      | Same as  <strong>CheckBox</strong>.                                                                        |
| <strong>ScrollBar</strong>         | An integer between the values specified for the  <strong>Max</strong> and <strong>Min</strong> properties. |
| <strong>SpinButton</strong>        | Same as  <strong>ScrollBar</strong>.                                                                       |
| <strong>ComboBox, ListBox</strong> | The value in the  <strong>BoundColumn</strong> of the currently selected rows.                             |
| <strong>CommandButton</strong>     | Always  <strong>False</strong>.                                                                            |
| <strong>MultiPage</strong>         | An integer indicating the currently active page.                                                           |
|                                    | Zero (0) indicates the first page. The maximum value is one less than the number of pages.                 |
| <strong>TextBox</strong>           | The text in the edit region.                                                                               |

 **Remarks**
 **BoundValue** applies to the control that has the focus.
The contents of the  **BoundValue** and **Value** properties are identical most of the time. When the user edits a control so that its value changes, the contents of **BoundValue** and **Value** are different until the change is final.
Several things occur when the user changes the value of a control. For example, if a user changes the text in a  **TextBox**, the following things occur:


1. The  **Change** event is initiated. At this time the **Value** property contains the new text and **BoundValue** contains the previous text.

2. The  **BeforeUpdate** event is initiated.

3. The  **AfterUpdate** event is initiated. The values for **BoundValue** and **Value** are once again identical, containing the new text.


 **BoundValue** cannot be used with a multi-select list box.

