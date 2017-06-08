---
title: NavigationButton.PressedForeTint Property (Access)
keywords: vbaac10.chm14625
f1_keywords:
- vbaac10.chm14625
ms.prod: access
api_name:
- Access.NavigationButton.PressedForeTint
ms.assetid: 70267cd4-ed42-9533-4cb6-e4338fa38fc1
ms.date: 06/08/2017
---


# NavigationButton.PressedForeTint Property (Access)

Gets or sets the tint applied to the theme color in the  **PressedForeColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **PressedForeTint**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

The  **PressedForeTint** property contains a numeric expression that can be used to lighten the theme color in the **PressedForeColor** property. The default value of the **PressedForeTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)

