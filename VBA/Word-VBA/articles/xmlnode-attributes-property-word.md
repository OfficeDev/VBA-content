---
title: XMLNode.Attributes Property (Word)
keywords: vbawd10.chm37748751
f1_keywords:
- vbawd10.chm37748751
ms.prod: word
api_name:
- Word.XMLNode.Attributes
ms.assetid: 64731b03-12cb-1f48-30f5-0a1c5329ac47
ms.date: 06/08/2017
---


# XMLNode.Attributes Property (Word)

Returns an  **XMLNodes** collection that represents the attributes for the specified element.


## Syntax

 _expression_ . **Attributes**

 _expression_ Required. A variable that represents a **[XMLNode](xmlnode-object-word.md)** object.


## Remarks

All  **XMLNode** objects in the **XMLNodes** collection returned by using the **Attributes** property have a **NodeType** property value of **wdXMLNodeAttribute** .


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

