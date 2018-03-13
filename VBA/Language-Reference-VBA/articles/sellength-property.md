---
title: SelLength Property
keywords: fm20.chm2001870
f1_keywords:
- fm20.chm2001870
ms.prod: office
api_name:
- Office.SelLength
ms.assetid: 86f86e84-b22e-a86a-12b9-dc1011cbcf9d
ms.date: 06/08/2017
---


# SelLength Property



The number of characters selected in a text box or the text portion of a combo box.
 **Syntax**
 _object_. **SelLength** [= _Long_ ]
The  **SelLength** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                |
|:----------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                                                                                                                                                                                                                                   |
| <em>Long</em>         | Optional. A numeric expression specifying the number of characters selected. For  <strong>SelLength</strong> and <strong>SelStart</strong>, the valid range of settings is 0 to the total number of characters in the edit area of a <strong>ComboBox</strong> or <strong>TextBox</strong>. |

 **Remarks**
The  **SelLength** property is always valid, even when the control does not have[focus](vbe-glossary.md). Setting  **SelLength** to a value less than zero creates an error. Attempting to set **SelLength** to a value greater than the number of characters available in a control results in a value equal to the number of characters in the control.

 **Note**  Changing the value of the  **SelStart** property cancels any existing selection in the control, places an insertion point in the text, and sets **SelLength** to zero.

The default value, zero, means that no text is currently selected.

