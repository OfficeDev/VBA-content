---
title: XMLNode.BaseName Property (Word)
keywords: vbawd10.chm37748736
f1_keywords:
- vbawd10.chm37748736
ms.prod: word
api_name:
- Word.XMLNode.BaseName
ms.assetid: 770e276b-8bf5-9f0d-64bd-e7df29a71233
ms.date: 06/08/2017
---


# XMLNode.BaseName Property (Word)

Returns a  **String** that represents the name of the element without any prefix.


## Syntax

 _expression_ . **BaseName**

 _expression_ Required. A variable that represents a **[XMLNode](xmlnode-object-word.md)** object.


## Example

The following example adds the author attribute to the book element in the active document and then sets the value of the attribute.


```vb
Sub AddIDAttribute() 
 Dim objElement As XMLNode 
 Dim objAttribute As XMLNode 
 
 For Each objElement In ActiveDocument.XMLNodes 
 If objElement.NodeType = wdXMLNodeElement Then 
 If objElement.BaseName = "book" Then 
 
 Set objAttribute = objElement.Attributes _ 
 .Add("author", objElement.NamespaceURI) 
 
 objAttribute.NodeValue = "David Barber" 
 
 Exit For 
 End If 
 End If 
 Next 
End Sub
```


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

