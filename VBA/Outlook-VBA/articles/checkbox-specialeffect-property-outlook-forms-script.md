---
title: CheckBox.SpecialEffect Property (Outlook Forms Script)
keywords: olfm10.chm2001960
f1_keywords:
- olfm10.chm2001960
ms.prod: outlook
ms.assetid: 98b61ff8-95c9-42cb-aa71-25053f0e6644
ms.date: 06/08/2017
---


# CheckBox.SpecialEffect Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the visual appearance of an object. Read/write.


## Syntax

 _expression_. **SpecialEffect**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

The settings for  **SpecialEffect** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Object appears flat, distinguished from the surrounding form by a border, a change of color, or both.|
|2|Object has a shadow on the top and left and a highlight on the bottom and right. The control and its border appear to be carved into the form that contains them. Default for  **[CheckBox](checkbox-object-outlook-forms-script.md)**.|
 **SpecialEffect** uses the system colors to define its borders.


