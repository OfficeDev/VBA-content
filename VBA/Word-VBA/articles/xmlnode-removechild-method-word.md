---
title: XMLNode.RemoveChild Method (Word)
keywords: vbawd10.chm37748838
f1_keywords:
- vbawd10.chm37748838
ms.prod: word
api_name:
- Word.XMLNode.RemoveChild
ms.assetid: 9c4d0e0a-ab58-7c9f-9fc2-f07a28281c29
ms.date: 06/08/2017
---


# XMLNode.RemoveChild Method (Word)

Removes a child element from the specified element.


## Syntax

 _expression_ . **RemoveChild**( **_ChildElement_** )

 _expression_ An expression that returns an **XMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChildElement_|Required| **XMLNode**|The child element to be removed.|

### Return Value

Nothing


## Example

The following example removes the first child from the first element in the active document.


```vb
ActiveDocument.XMLNodes(1).RemoveChild _ 
 ActiveDocument.XMLNodes(1).ChildNodes(1)
```


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

