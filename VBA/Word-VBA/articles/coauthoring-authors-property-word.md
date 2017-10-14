---
title: CoAuthoring.Authors Property (Word)
keywords: vbawd10.chm254869505
f1_keywords:
- vbawd10.chm254869505
ms.prod: word
api_name:
- Word.CoAuthoring.Authors
ms.assetid: 95d7d241-505b-a282-1f20-4486149433ad
ms.date: 06/08/2017
---


# CoAuthoring.Authors Property (Word)

 Returns a **[CoAuthors](coauthors-object-word.md)** collection that represents all the co authors currently editing the document. Read-only.


## Syntax

 _expression_ . **Authors**

 _expression_ An expression that returns a **[CoAuthoring](coauthoring-object-word.md)** object.


## Remarks

The collection returned by this property is static. If this collection is stored and then new users begin editing the document, or current users are no longer editing the document, the stored collection will not change.


## Example

The following code example gets all the co authors currently editing the document.


```vb
Dim allAuthors As CoAuthors 
Set allAuthors = ActiveDocument.CoAuthoring.Authors
```


## See also


#### Concepts


[CoAuthoring Object](coauthoring-object-word.md)

