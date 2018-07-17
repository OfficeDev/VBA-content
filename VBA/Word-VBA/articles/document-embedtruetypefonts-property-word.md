---
title: Document.EmbedTrueTypeFonts Property (Word)
keywords: vbawd10.chm158007346
f1_keywords:
- vbawd10.chm158007346
ms.prod: word
api_name:
- Word.Document.EmbedTrueTypeFonts
ms.assetid: ac8fb6a1-584a-2ddb-4216-53e30473ff65
ms.date: 06/08/2017
---


# Document.EmbedTrueTypeFonts Property (Word)

 **True** if Microsoft Word embeds TrueType fonts in a document when it is saved. Read/write **Boolean** .


## Syntax

 _expression_ . **EmbedTrueTypeFonts**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Embedding TrueType fonts allows others to view a document with the same fonts that were used to create it. 


## Example

This example sets Word to automatically embed TrueType fonts when saving a document, and then it saves the active document.


```vb
ActiveDocument.EmbedTrueTypeFonts = True 
ActiveDocument.Save
```

This example returns the current status of the  **Embed TrueType** fonts check box in the **Save** options area on the **Save** tab in the **Options** dialog box.




```
temp = ActiveDocument.EmbedTrueTypeFonts
```


## See also


#### Concepts


[Document Object](document-object-word.md)

