---
title: Style Object (Word)
keywords: vbawd10.chm2348
f1_keywords:
- vbawd10.chm2348
ms.prod: word
api_name:
- Word.Style
ms.assetid: 473f8f41-2cba-769e-c0da-441d9d85b009
ms.date: 06/08/2017
---


# Style Object (Word)

Represents a single built-in or user-defined style. The  **Style** object includes style attributes (such as font, font style, and paragraph spacing) as properties of the **Style** object. The **Style** object is a member of the **Styles** collection. The **[Styles](styles-object-word.md)** collection includes all the styles in the specified document.


## Remarks

Use  **Styles** (Index), where Index is the style name, a **WdBuiltinStyle** constant or index number, to return a single **Style** object. You must exactly match the spelling and spacing of the style name, but not necessarily its capitalization. The following example modifies the font name of the user-defined style named "Color" in the active document.


```vb
ActiveDocument.Styles("Color").Font.Name = "Arial"
```

The following example sets the built-in Heading 1 style to not be bold.




```vb
ActiveDocument.Styles(wdStyleHeading1).Font.Bold = False
```

The style index number represents the position of the style in the alphabetically sorted list of style names. Note that  `Styles(1)` is the first style in the alphabetical list. The following example displays the base style and style name of the first style in the **[Styles](styles-object-word.md)** collection.




```vb
MsgBox "Base style= " _ 
 &; ActiveDocument.Styles(1).BaseStyle &; vbCr _ 
 &; "Style name= " &; ActiveDocument.Styles(1).NameLocal
```

To apply a style to a range, paragraph, or multiple paragraphs, set the  **Style** property to a user-defined or built-in style name. The following example applies the Normal style to the first four paragraphs in the active document.




```vb
Set myRange = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(1).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End) 
myRange.Style = wdStyleNormal
```

The following example applies the Heading 1 style to the first paragraph in the selection.




```
Selection.Paragraphs(1).Style = wdStyleHeading1
```

The following example creates a character style named "Bolded" and applies it to the selection.




```vb
Set myStyle = ActiveDocument.Styles.Add(Name:="Bolded", _ 
 Type:=wdStyleTypeCharacter) 
myStyle.Font.Bold = True 
Selection.Range.Style = "Bolded"
```

Use the  **OrganizerCopy** method to copy styles between documents and templates. Use the **UpdateStyles** method to update the styles in the active document to match the style definitions in the attached template. Use the **OpenAsDocument** method to open a template as a document so that you can modify the template styles.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


