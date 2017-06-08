---
title: Document.CurrentRsid Property (Word)
keywords: vbawd10.chm158007859
f1_keywords:
- vbawd10.chm158007859
ms.prod: word
api_name:
- Word.Document.CurrentRsid
ms.assetid: 500a743e-6d1e-e93d-b4d2-20ac13c4651a
ms.date: 06/08/2017
---


# Document.CurrentRsid Property (Word)

Returns a  **Long** that represents a random number that Word assigns to changes in a document. Read-only.


## Syntax

 _expression_ . **CurrentRsid**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

If the  **[StoreRSIDOnSave](options-storersidonsave-property-word.md)** property is **True** , each time a document is saved, Word generates a random number that the application uses to facilitate comparing and merging documents. Word stores the random numbers in a table and updates the table after each save. The **CurrentRsid** property returns the last number that Word has assigned to a document.


## See also


#### Concepts


[Document Object](document-object-word.md)

