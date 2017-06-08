---
title: Paragraph.OutlineLevel Property (Word)
keywords: vbawd10.chm156696778
f1_keywords:
- vbawd10.chm156696778
ms.prod: word
api_name:
- Word.Paragraph.OutlineLevel
ms.assetid: 657141b2-c02c-b3f5-5cf3-f92c5720bb28
ms.date: 06/08/2017
---


# Paragraph.OutlineLevel Property (Word)

Returns or sets the outline level for the specified paragraph. Read/write  **[WdOutlineLevel](wdoutlinelevel-enumeration-word.md)** .


## Syntax

 _expression_ . **OutlineLevel**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

If a paragraph has a heading style applied to it (Heading 1 through Heading 9), the outline level is the same as the heading style and cannot be changed. Outline levels are visible only in outline view or the document map pane.


## Example

This example returns the outline level of the first paragraph in the active document.


```
temp = ActiveDocument.Paragraphs(1).OutlineLevel
```

This example sets the outline level for each paragraph in the active document. First the Normal style is applied to all paragraphs. The  **Mod** operator is used to determine which outline level (1, 2, or 3) to apply to successive paragraphs in the document, and then the view is changed to outline view.




```vb
Set myParas = ActiveDocument.Paragraphs 
ActiveDocument.Paragraphs.Style = wdStyleNormal 
For x = 1 To myParas.Count 
 If x Mod 3 = 1 Then 
 myParas(x).OutlineLevel = wdOutlineLevel1 
 ElseIf x Mod 3 = 2 Then 
 myParas(x).OutlineLevel = wdOutlineLevel2 
 Else 
 myParas(x).OutlineLevel = wdOutlineLevel3 
 End If 
Next x 
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

