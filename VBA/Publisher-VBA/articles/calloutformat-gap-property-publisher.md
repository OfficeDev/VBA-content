---
title: CalloutFormat.Gap Property (Publisher)
keywords: vbapb10.chm2490631
f1_keywords:
- vbapb10.chm2490631
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Gap
ms.assetid: fd7cdac7-5f09-a574-e9ef-08feebd81cff
ms.date: 06/08/2017
---


# CalloutFormat.Gap Property (Publisher)

Returns or sets a  **Variant** indicating the horizontal distance between the end of the callout line and the text bounding box. Read/write.


## Syntax

 _expression_. **Gap**

 _expression_A variable that represents a  **CalloutFormat** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

This example sets the distance between the callout line and the text bounding box to 3 points for the first shape in the active publication. For the example to work, the shape must be a callout.


```vb
ActiveDocument.Pages(1).Shapes(1).Callout.Gap = 3
```


