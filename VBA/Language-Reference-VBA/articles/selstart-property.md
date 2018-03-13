---
title: SelStart Property
keywords: fm20.chm5225091
f1_keywords:
- fm20.chm5225091
ms.prod: office
api_name:
- Office.SelStart
ms.assetid: ca0db01c-bea0-6827-376f-f2a41c4eb5ed
ms.date: 06/08/2017
---


# SelStart Property



Indicates the starting point of selected text, or the insertion point if no text is selected.
 **Syntax**
 _object_. **SelStart** [= _Long_ ]
The  **SelStart** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                             |
|:----------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                                                                                                                                                                                                                                                                |
| <em>Long</em>         | Optional. A numeric expression specifying the starting point of text selected. For  <strong>SelLength</strong> and <strong>SelStart</strong>, the valid range of settings is 0 to the total number of characters in the edit area of a <strong>ComboBox</strong> or <strong>TextBox</strong>. The default value is zero. |

 **Remarks**
The  **SelStart** property is always valid, even when the control does not have[focus](vbe-glossary.md). Setting  **SelStart** to a value less than zero creates an error. Attempting to set **SelStart** to a value greater than the number of characters available in a control results in a value equal to the number of characters in the control.
Changing the value of  **SelStart** cancels any existing selection in the control, places an insertion point in the text, and sets the **SelLength** property to zero.

