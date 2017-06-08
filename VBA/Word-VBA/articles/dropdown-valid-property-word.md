---
title: DropDown.Valid Property (Word)
keywords: vbawd10.chm153419776
f1_keywords:
- vbawd10.chm153419776
ms.prod: word
api_name:
- Word.DropDown.Valid
ms.assetid: 2ac906a1-effa-02ff-85db-620f30434e89
ms.date: 06/08/2017
---


# DropDown.Valid Property (Word)

 **True** if the specified form field object is a valid drop down form field. Read-only **Boolean** .


## Syntax

 _expression_ . **Valid**

 _expression_ Required. An expression that returns a **[DropDown](dropdown-object-word.md)** object.


## Remarks

Use the  **Type** property of the **[FormField](formfield-object-word.md)** object to determine the type of form field ( **wdFieldFormDropDown** ) before applying the **[DropDown](formfield-dropdown-property-word.md)** property. This precaution ensures that the **FormField** object is of the expected type.


## See also


#### Concepts


[DropDown Object](dropdown-object-word.md)

