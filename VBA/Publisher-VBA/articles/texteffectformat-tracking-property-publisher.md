---
title: TextEffectFormat.Tracking Property (Publisher)
keywords: vbapb10.chm3735825
f1_keywords:
- vbapb10.chm3735825
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.Tracking
ms.assetid: 9e110e21-be0c-ec49-6bc4-1ff210de141c
ms.date: 06/08/2017
---


# TextEffectFormat.Tracking Property (Publisher)

Returns or sets a  **Variant** indicating the tracking value used to display space between the characters in the specified text range. Read/write.


## Syntax

 _expression_. **Tracking**

 _expression_A variable that represents a  **TextEffectFormat** object.


## Remarks

Valid range is a  **float** value between 0.0 and 5.0 points. Setting the property to 0.0 disables tracking. Indeterminate values are returned as -2.


## Example

This example disables tracking in the second story by setting the  **Tracking** property to zero.


```vb
Sub DisableTracking() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Tracking = 0.0 
 
End Sub
```


