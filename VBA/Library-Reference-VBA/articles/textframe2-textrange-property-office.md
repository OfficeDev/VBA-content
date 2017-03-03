---
title: TextFrame2.TextRange Property (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.TextFrame2.TextRange
ms.assetid: 6ea3de69-5c3d-2f54-c8c6-df80dab8fa62
---


# TextFrame2.TextRange Property (Office)

Sets the text for a range of nodes in a SmartArt object. Read-only


## Syntax

 _expression_. **TextRange**

 _expression_ An expression that returns a **TextFrame2** object.


## Example

The following example sets the text inside the first node.


```
smartart.AllNodes(1).TextFrame2.TextRange.Text="Node 1"
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-office.md)

