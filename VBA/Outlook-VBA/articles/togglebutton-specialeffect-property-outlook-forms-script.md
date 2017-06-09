---
title: ToggleButton.SpecialEffect Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: a123389c-3bf4-843f-763c-76e8fff18c6c
ms.date: 06/08/2017
---


# ToggleButton.SpecialEffect Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the visual appearance of an object. Read/write.


## Syntax

 _expression_. **SpecialEffect**

 _expression_A variable that represents a  **ToggleButton** object.


## Remarks

The settings for SpecialEffect are:



|**Value**|**Description**|
|:-----|:-----|
|0|Object appears flat, distinguished from the surrounding form by a border, a change of color, or both.|
|2|Object has a shadow on the top and left and a highlight on the bottom and right. The control and its border appear to be carved into the form that contains them. Default for  **[ToggleButton](togglebutton-object-outlook-forms-script.md)**.|
 **SpecialEffect** uses the system colors to define its borders.

Although the  **SpecialEffect** property exists on the **ToggleButton**, the property is disabled. You cannot set or return a value for this property on the  **ToggleButton** **ToggleButton**.


