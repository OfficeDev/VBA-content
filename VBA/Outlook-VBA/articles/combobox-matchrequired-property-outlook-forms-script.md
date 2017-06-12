---
title: ComboBox.MatchRequired Property (Outlook Forms Script)
keywords: olfm10.chm2001500
f1_keywords:
- olfm10.chm2001500
ms.prod: outlook
ms.assetid: 01d6c98b-ab87-d968-011b-7acfa2058feb
ms.date: 06/08/2017
---


# ComboBox.MatchRequired Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a value entered in the text portion of a **[ComboBox](combobox-object-outlook-forms-script.md)** must match an entry in the existing list portion of the control. Read/write.


## Syntax

 _expression_. **MatchRequired**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The user can enter non-matching values, but may not leave the control until a matching value is entered.

 **True** if the text entered must match an existing list entry. **False** if the text entered can be different from all existing list entries (default).

If the  **MatchRequired** property is **True**, the user cannot exit the  **ComboBox** until the text entered matches an entry in the existing list. **MatchRequired** maintains the integrity of the list by requiring the user to select an existing entry.

Not all containers enforce this property.


