---
title: OlkComboBox.SelLength Property (Outlook)
keywords: vbaol11.chm1000222
f1_keywords:
- vbaol11.chm1000222
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.SelLength
ms.assetid: 3cbd5016-3868-6cf9-c28c-8d692620f367
ms.date: 06/08/2017
---


# OlkComboBox.SelLength Property (Outlook)

Returns or sets a  **Long** that specifies the number of characters in the current selection. Read/write.


## Syntax

 _expression_ . **SelLength**

 _expression_ A variable that represents an **OlkComboBox** object.


## Remarks

The current selection is specified by  **[SelText](olkcombobox-seltext-property-outlook.md)** , which is a portion of the control's **[Value](olkcombobox-value-property-outlook.md)** . The maximum number of characters that can be supported for **Value** is **[MaxLength](olkcombobox-maxlength-property-outlook.md)** .

The default value is zero, which means no text is currently selected.

The  **SelLength** property is always valid, even when the control does not have focus.

Setting  **SelLength** to a value less than zero causes an error. Attempting to set the value greater than **MaxLength** results in setting **SelLength** to **MaxLength** .


## See also


#### Concepts


[OlkComboBox Object](olkcombobox-object-outlook.md)

