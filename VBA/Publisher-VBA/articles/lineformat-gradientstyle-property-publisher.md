---
title: LineFormat.GradientStyle Property (Publisher)
keywords: vbapb10.chm3408151
f1_keywords:
- vbapb10.chm3408151
ms.prod: publisher
ms.assetid: e5416db9-a145-8f71-2d75-1720191922bb
ms.date: 06/08/2017
---


# LineFormat.GradientStyle Property (Publisher)

Returns the gradient style for the specified line. Read-only.


## Syntax

 _expression_. **GradientStyle**

 _expression_A variable that represents a  **LineFormat** object.


## Return value

 **MsoGradientStyle**


## Remarks

Attempting to return this property for a line that doesn't have a gradient generates an error. Use the  **[Type](lineformat-type-property-publisher.md)** property to determine whether the line has a gradient.


## See also


#### Concepts


 [LineFormat Object](lineformat-object-publisher.md)

