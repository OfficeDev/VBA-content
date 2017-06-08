---
title: OlkTextBox.SelLength Property (Outlook)
keywords: vbaol11.chm1000063
f1_keywords:
- vbaol11.chm1000063
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.SelLength
ms.assetid: 89d040ba-b28f-20f1-e449-1c533370b711
ms.date: 06/08/2017
---


# OlkTextBox.SelLength Property (Outlook)

Returns or sets a  **Long** that specifies the number of characters in the current selection. Read/write.


## Syntax

 _expression_ . **SelLength**

 _expression_ A variable that represents an **OlkTextBox** object.


## Remarks

The current selection is specified by  **[SelText](olktextbox-seltext-property-outlook.md)** , which is a portion of the control's **[Value](olktextbox-value-property-outlook.md)** . The maximum number of characters that can be supported for **Value** is **[MaxLength](olktextbox-maxlength-property-outlook.md)** .

The default value is zero, which means no text is currently selected.

The  **SelLength** property is always valid, even when the control does not have focus.

Setting  **SelLength** to a value less than zero causes an error. Attempting to set the value greater than **MaxLength** results in setting **SelLength** to **MaxLength** .


## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

