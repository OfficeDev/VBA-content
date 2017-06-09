---
title: OlkTextBox.SelText Property (Outlook)
keywords: vbaol11.chm1000064
f1_keywords:
- vbaol11.chm1000064
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.SelText
ms.assetid: ba529e92-8a28-1c50-bf0a-0e67ae3645bc
ms.date: 06/08/2017
---


# OlkTextBox.SelText Property (Outlook)

Returns a  **String** that represents the currently selected portion of the value of the text box. Read-only.


## Syntax

 _expression_ . **SelText**

 _expression_ A variable that represents an **OlkTextBox** object.


## Remarks

 **SelText** represents the current selection, which is a portion of the control's **[Value](olktextbox-value-property-outlook.md)** . The maximum number of characters that can be supported for **Value** is **[MaxLength](olktextbox-maxlength-property-outlook.md)** .

The default value is an empty string.


## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

