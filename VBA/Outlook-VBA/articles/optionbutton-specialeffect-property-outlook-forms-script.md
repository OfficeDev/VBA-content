---
title: OptionButton.SpecialEffect Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 908f588a-8f76-82d8-8b0f-1cb7764b179c
ms.date: 06/08/2017
---


# OptionButton.SpecialEffect Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the visual appearance of an object. Read/write.


## Syntax

 _expression_. **SpecialEffect**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks

The settings for  **SpecialEffect** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Object appears flat, distinguished from the surrounding form by a border, a change of color, or both.|
|2|Object has a shadow on the top and left and a highlight on the bottom and right. The control and its border appear to be carved into the form that contains them. Default for  **[OptionButton](optionbutton-object-outlook-forms-script.md)**, valid for all controls (default).|
 **SpecialEffect** uses the system colors to define its borders.


