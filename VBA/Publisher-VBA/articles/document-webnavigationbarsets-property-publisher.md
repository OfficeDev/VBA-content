---
title: Document.WebNavigationBarSets Property (Publisher)
keywords: vbapb10.chm196741
f1_keywords:
- vbapb10.chm196741
ms.prod: publisher
api_name:
- Publisher.Document.WebNavigationBarSets
ms.assetid: 4193dbce-a2e3-2587-5282-43b4c3cec921
ms.date: 06/08/2017
---


# Document.WebNavigationBarSets Property (Publisher)

Returns a  **WebNavigationBarSets** object representing a collection of all **WebNavigationBarSet** objects in the specified document. Read-only.


## Syntax

 _expression_. **WebNavigationBarSets**

 _expression_A variable that represents a  **Document** object.


### Return Value

WebNavigationBarSets


## Example

The following example sets an object variable to the collection of Web navigation bar sets in the active document and adds a new navigation bar set to it.


```vb
Dim objWebNavBarSets As WebNavigationBarSets 
 
Set objWebNavBarSets = ActiveDocument.WebNavigationBarSets 
objWebNavBarSets.AddSet _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignBracket, _ 
 AutoUpdate:=True 

```


