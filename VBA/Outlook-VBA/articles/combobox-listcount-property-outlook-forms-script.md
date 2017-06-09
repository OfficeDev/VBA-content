---
title: ComboBox.ListCount Property (Outlook Forms Script)
keywords: olfm10.chm2001410
f1_keywords:
- olfm10.chm2001410
ms.prod: outlook
ms.assetid: 8ea1e997-470f-1336-5a72-ce66ece1f292
ms.date: 06/08/2017
---


# ComboBox.ListCount Property (Outlook Forms Script)

Returns a  **Long** that represents the number of list entries in a control. Read-only.


## Syntax

 _expression_. **ListCount**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

 **ListCount** is the number of rows over which you can scroll. **ListCount** is always one greater than the largest value for the **[ListIndex](combobox-listindex-property-outlook-forms-script.md)** property, because index numbers begin with 0 and the count of items begins with 1. If no item is selected, **ListCount** is 0 and **ListIndex** is -1.


