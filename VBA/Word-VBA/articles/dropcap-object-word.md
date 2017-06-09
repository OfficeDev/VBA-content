---
title: DropCap Object (Word)
keywords: vbawd10.chm2390
f1_keywords:
- vbawd10.chm2390
ms.prod: word
api_name:
- Word.DropCap
ms.assetid: 79daea90-657b-43db-34e3-08f7aed74591
ms.date: 06/08/2017
---


# DropCap Object (Word)

Represents a dropped capital letter at the beginning of a paragraph. There is no  **DropCaps** collection; each **[Paragraph](paragraph-object-word.md)** object contains only one **DropCap** object.


## Remarks

Use the  **DropCap** property to return a **DropCap** object. The following example sets a dropped capital letter for the first letter in the first paragraph in the active document.


```vb
With ActiveDocument.Paragraphs(1).DropCap 
 .Enable 
 .Position = wdDropNormal 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

