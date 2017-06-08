---
title: XMLNode.ChildNodes Property (Word)
keywords: vbawd10.chm37748749
f1_keywords:
- vbawd10.chm37748749
ms.prod: word
api_name:
- Word.XMLNode.ChildNodes
ms.assetid: 79d5e434-be1a-6420-ac82-ecf9c7c49e32
ms.date: 06/08/2017
---


# XMLNode.ChildNodes Property (Word)

Returns an  **XMLNodes** collection that represents the child elements of a specified element.


## Syntax

 _expression_ . **ChildNodes**

 _expression_ Required. A variable that represents a **[XMLNode](xmlnode-object-word.md)** object.


## Example

The following example removes the first child element of the root element in the active document.


```vb
ActiveDocument.XMLNodes(1).RemoveChild _ 
 ActiveDocument.XMLNodes(1).ChildNodes(1)
```


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

