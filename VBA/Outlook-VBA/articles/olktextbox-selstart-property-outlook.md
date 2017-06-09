---
title: OlkTextBox.SelStart Property (Outlook)
keywords: vbaol11.chm1000062
f1_keywords:
- vbaol11.chm1000062
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.SelStart
ms.assetid: cca8ffc2-4c68-72f5-7e09-6f8845d72e35
ms.date: 06/08/2017
---


# OlkTextBox.SelStart Property (Outlook)

Returns or sets a  **Long** that specifies either the starting point of the selected text or the insertion point if no text has been selected. Read/write.


## Syntax

 _expression_ . **SelStart**

 _expression_ A variable that represents an **OlkTextBox** object.


## Remarks

The current selection is specified by  **[SelText](olktextbox-seltext-property-outlook.md)** , which is a portion of the control's **[Value](olktextbox-value-property-outlook.md)** . The maximum number of characters that can be supported for **Value** is **[MaxLength](olktextbox-maxlength-property-outlook.md)** .

The default value is zero, which means no text is selected and the insertion point is at the beginning.

The  **SelStart** property is always valid, even when the control does not have focus. Setting **SelStart** to a value less than zero causes an error. Setting **SelStart** to a value greater than **MaxLength** will reset **SelStart** to **MaxLength** . Changing the value of **SelStart** cancels any existing selection, places the insertion point in the text, and sets the **SelLength** property to zero.


## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

