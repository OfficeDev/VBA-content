---
title: ComboBox.MatchFound Property (Outlook Forms Script)
keywords: olfm10.chm2001490
f1_keywords:
- olfm10.chm2001490
ms.prod: outlook
ms.assetid: 2e35541f-990d-fa2a-4431-695f9d951c98
ms.date: 06/08/2017
---


# ComboBox.MatchFound Property (Outlook Forms Script)

Returns a  **Boolean** value that indicates whether the text that a user has typed into a **[ComboBox](combobox-object-outlook-forms-script.md)** matches any of the entries in the list. Read-only.


## Syntax

 _expression_. **MatchFound**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

 **True** if the contents of the **[Value](combobox-value-property-outlook-forms-script.md)** property matches one of the records in the list. **False** if the contents of **Value** does not match any of the records in the list (default).

The  **MatchFound** property is read-only. It is not applicable when the **[MatchEntry](combobox-matchentry-property-outlook-forms-script.md)** property is set to 2.


