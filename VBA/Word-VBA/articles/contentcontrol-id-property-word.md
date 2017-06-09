---
title: ContentControl.ID Property (Word)
keywords: vbawd10.chm266534930
f1_keywords:
- vbawd10.chm266534930
ms.prod: word
api_name:
- Word.ContentControl.ID
ms.assetid: 2a9480f0-c572-6724-121f-b1a41d99ab93
ms.date: 06/08/2017
---


# ContentControl.ID Property (Word)

Returns a  **String** that represents the identification for a content control. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

The  **ID** property is an internal number that you cannot change but that you can use to identify a content control in code. This number is unique for each content control and does not change.

When you get the  **ID** property value at runtime, it is returned as an unsigned value. However, when saved into the Office Open XML file format, it is saved as a signed value. If your solution attempts to map programmatically returned values to values saved in the file format, you must check for both the unsigned and signed version of the value obtained from this property.


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

