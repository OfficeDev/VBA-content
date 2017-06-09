---
title: Document.ReadabilityStatistics Property (Word)
keywords: vbawd10.chm158007392
f1_keywords:
- vbawd10.chm158007392
ms.prod: word
api_name:
- Word.Document.ReadabilityStatistics
ms.assetid: e9da9d92-bc1f-d575-07b1-3eae2749a9e5
ms.date: 06/08/2017
---


# Document.ReadabilityStatistics Property (Word)

Returns a  **ReadabilityStatistics** collection that represents the readability statistics for the specified document or range. Read-only.


## Syntax

 _expression_ . **ReadabilityStatistics**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays each readability statistic, along with its value, for document one.


```vb
For Each rs In Documents(1).ReadabilityStatistics 
 Msgbox rs.Name &; " - " &; rs.Value 
Next rs
```


## See also


#### Concepts


[Document Object](document-object-word.md)

