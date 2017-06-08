---
title: ListParagraphs Object (Word)
ms.prod: word
ms.assetid: 759c510b-bca1-0b4b-005c-5a3783dd8e96
ms.date: 06/08/2017
---


# ListParagraphs Object (Word)

A collection of  **Paragraph** objects that represents the paragraphs of the specified document, list, or range that have list formatting applied.


## Remarks

Use the  **ListParagraphs** property to return the **ListParagraphs** collection. The following example applies highlighting to the collection of paragraphs with list formatting in the active document.


```vb
For Each para in ActiveDocument.ListParagraphs 
 para.Range.HighlightColorIndex = wdTurquoise 
Next para
```

Use  **ListParagraphs** (Index), where Index is the index number, to return a single **[Paragraph](paragraph-object-word.md)** object with list formatting.

Paragraphs can have two types of list formatting. The first type includes an automatically added number or bullet at the beginning of each paragraph in the list. The second type includes LISTNUM fields, which can be placed anywhere inside a paragraph. There can be more than one LISTNUM field per paragraph.

To add list formatting to paragraphs, you can use the  **ApplyListTemplate** , **ApplyBulletDefault** , **ApplyNumberDefault** , or **ApplyOutlineNumberDefault** method. You access these methods through the **ListFormat** object for a specified range.

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


