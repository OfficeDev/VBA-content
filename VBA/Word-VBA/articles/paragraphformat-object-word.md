---
title: ParagraphFormat Object (Word)
keywords: vbawd10.chm2387
f1_keywords:
- vbawd10.chm2387
ms.prod: word
api_name:
- Word.ParagraphFormat
ms.assetid: 712d754a-dc92-f1a3-531d-dfae74a42c23
ms.date: 06/08/2017
---


# ParagraphFormat Object (Word)

Represents all the formatting for a paragraph.


## Remarks

Use the  **Format** property to return the **ParagraphFormat** object for a paragraph or paragraphs. The **ParagraphFormat** property returns the **ParagraphFormat** object for a selection, range, style, **Find** object, or **Replacement** object. The following example centers the third paragraph in the active document.


```vb
ActiveDocument.Paragraphs(3).Format.Alignment = _ 
 wdAlignParagraphCenter
```

The following example finds the next double-spaced paragraph after the selection.




```vb
With Selection.Find 
 .ClearFormatting 
 .ParagraphFormat.LineSpacingRule = wdLineSpaceDouble 
 .Text = "" 
 .Forward = True 
 .Wrap = wdFindContinue 
End With 
Selection.Find.Execute
```

You can use Visual Basic's  **New** keyword to create a new, standalone **ParagraphFormat** object. The following example creates a **ParagraphFormat** object, sets some formatting properties for it, and then applies all of its properties to the first paragraph in the active document.




```vb
Dim myParaF As New ParagraphFormat 
myParaF.Alignment = wdAlignParagraphCenter 
myParaF.Borders.Enable = True 
ActiveDocument.Paragraphs(1).Format = myParaF
```

You can also make a standalone copy of an existing  **ParagraphFormat** object by using the **Duplicate** property. The following example duplicates the paragraph formatting of the first paragraph in the active document and stores the formatting in _myDup_ . The example changes the left indent of _myDup_ to 1 inch, creates a new document, inserts text into the document, and applies the paragraph formatting of _myDup_ to the text.




```vb
Set myDup = ActiveDocument.Paragraphs(1).Format.Duplicate 
myDup.LeftIndent = InchesToPoints(1) 
Documents.Add 
Selection.InsertAfter "This is a new paragraph." 
Selection.Paragraphs.Format = myDup
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

