---
title: TextFrame.ValidLinkTarget Method (Publisher)
keywords: vbapb10.chm3866662
f1_keywords:
- vbapb10.chm3866662
ms.prod: publisher
api_name:
- Publisher.TextFrame.ValidLinkTarget
ms.assetid: ee946f58-669f-7150-0f40-2dd3b857e274
ms.date: 06/08/2017
---


# TextFrame.ValidLinkTarget Method (Publisher)

Determines whether the text frame of one shape can be linked to the text frame of another shape. Returns  **True** if **_LinkTarget_** is a valid target, **False** if **_LinkTarget_** already contains text or is already linked, or if the shape does not support attached text.


## Syntax

 _expression_. **ValidLinkTarget**( **_LinkTarget_**)

 _expression_A variable that represents a  **TextFrame** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LinkTarget|Required| **Shape**|The shape with the target text frame to which you want to link the text frame returned by expression.|

### Return Value

Boolean


## Example

This example checks to see whether the text frames for the first and second shapes on the first page of the active publication can be linked to one another. If so, the example links the two text frames.


```vb
Dim txtFrame1 As TextFrame 
Dim txtFrame2 As TextFrame 
 
With ActiveDocument.Pages(1) 
 Set txtFrame1 = .Shapes(1).TextFrame 
 Set txtFrame2 = .Shapes(2).TextFrame 
End With 
 
If txtFrame1.ValidLinkTarget(LinkTarget:=txtFrame2.Parent) = True Then 
 txtFrame1.NextLinkedTextFrame = txtFrame2 
End If
```


