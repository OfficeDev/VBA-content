---
title: Font.Tracking Property (Publisher)
keywords: vbapb10.chm5373984
f1_keywords:
- vbapb10.chm5373984
ms.prod: publisher
api_name:
- Publisher.Font.Tracking
ms.assetid: c703a5ec-e8d7-36ce-ac50-d41265ce92db
ms.date: 06/08/2017
---


# Font.Tracking Property (Publisher)

Returns or sets a  **Variant** indicating the tracking value used to display space between the characters in the specified text range. Read/write.


## Syntax

 _expression_. **Tracking**

 _expression_A variable that represents a  **Font** object.


## Remarks

Valid range is 0.0 to 600.0 points. Setting the property to 0.0 disables tracking. Indeterminate values are returned as -2.


## Example

This example disables tracking in the second story by setting the  **Tracking** property to zero.


```vb
Sub DisableTracking() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Tracking = 0.0 
 
End Sub
```


