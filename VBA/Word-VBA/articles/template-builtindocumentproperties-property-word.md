---
title: Template.BuiltInDocumentProperties Property (Word)
keywords: vbawd10.chm157941768
f1_keywords:
- vbawd10.chm157941768
ms.prod: word
api_name:
- Word.Template.BuiltInDocumentProperties
ms.assetid: 48de083f-c24d-3991-e4dc-1e1acfdf1106
ms.date: 06/08/2017
---


# Template.BuiltInDocumentProperties Property (Word)

Returns a  **DocumentProperties** collection that represents all the built-in document properties for the specified document.


## Syntax

 _expression_ . **BuiltInDocumentProperties**

 _expression_ Required. A variable that represents a **[Template](template-object-word.md)** object.


## Remarks

To return a single  **DocumentProperty** object that represents a specific built-in document property, use the **BuiltinDocumentProperties** property. If Microsoft Word doesn't define a value for one of the built-in document properties, reading the **Value** property for that document property generates an error.

Use the  **CustomDocumentProperties** property to return the collection of custom document properties.

 For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## See also


#### Concepts


[Template Object](template-object-word.md)

