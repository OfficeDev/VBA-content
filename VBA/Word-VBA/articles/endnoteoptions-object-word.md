---
title: EndnoteOptions Object (Word)
ms.prod: word
api_name:
- Word.EndnoteOptions
ms.assetid: b63cf439-2297-fec9-ba36-66ad3f43dcbc
ms.date: 06/08/2017
---


# EndnoteOptions Object (Word)

Represents the properties assigned to a range or selection of endnotes in a document.


## Remarks

Use the  **EndnoteOptions** property of the **[Range](range-object-word.md)** or **[Selection](selection-object-word.md)** object to return an **EndnoteOptions** object.

Using the  **EndnoteOptions** object, you can assign different endnote properties to different areas of a document. For example, you may want endnotes in the introduction of a long document to be displayed as lowercase Roman numerals, while in the rest of your document they are displayed as Arabic numerals. The following example uses the **[NumberingRule](endnoteoptions-numberingrule-property-word.md)** , **[NumberStyle](endnoteoptions-numberstyle-property-word.md)** , and **[StartingNumber](endnoteoptions-startingnumber-property-word.md)** properties to format the endnotes in the first section ofthe active document.




```vb
Sub BookIntro() 
 Dim rngIntro As Range 
 
 'Sets the range as section one of the active document 
 Set rngIntro = ActiveDocument.Sections(1).Range 
 
 'Formats the EndnoteOptions properties 
 With rngIntro.EndnoteOptions 
 .NumberingRule = wdRestartSection 
 .NumberStyle = wdNoteNumberStyleLowercaseRoman 
 .StartingNumber = 1 
 End With 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


