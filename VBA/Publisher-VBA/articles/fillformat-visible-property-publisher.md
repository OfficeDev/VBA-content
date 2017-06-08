---
title: FillFormat.Visible Property (Publisher)
keywords: vbapb10.chm2359571
f1_keywords:
- vbapb10.chm2359571
ms.prod: publisher
api_name:
- Publisher.FillFormat.Visible
ms.assetid: 9cbb2604-6c33-de51-71f4-8c0304868cb5
ms.date: 06/08/2017
---


# FillFormat.Visible Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether the specified object or the formatting applied to the specified object is visible. Read/write.


## Syntax

 _expression_. **Visible**

 _expression_A variable that represents a  **FillFormat** object.


## Remarks

The  **Visible** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The specified object or formatting is not visible.|
| **msoTriStateMixed**|Return value only. The specified shape range contains both objects with visible formatting and objects with invisible formatting.|
| **msoTriStateToggle**| Set value only. Switches the specified object between visible and invisble.|
| **msoTrue**|The specified object or formatting is visible.|

## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on the first page in the active publication. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape does not already have a shadow, this example adds one to it.


```vb
With ActiveDocument.Pages(1).Shapes(3).Shadow 
 .Visible = msoTrue 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```


