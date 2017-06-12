---
title: Applying Formatting to Text
keywords: vbawd10.chm5209915
f1_keywords:
- vbawd10.chm5209915
ms.prod: word
ms.assetid: c20713bb-0e67-01d4-c9d4-91415658c0d7
ms.date: 06/08/2017
---


# Applying Formatting to Text

This topic includes Visual Basic examples related to the following tasks:


-  [Applying formatting to the selection](#Selection)
    
-  [Applying formatting to a range](#Range)
    
-  [Inserting text and applying character and paragraph formatting](#Inserting)
    
-  [Switching the space before a paragraph between 12 points and none](#TogglingSpace)
    
-  [Switching bold formatting on and off](#TogglingBold)
    
-  [Increasing the left margin by 0.5 inch](#Increasing)
    

## Applying formatting to the selection

The following example uses the  **[Selection](application-selection-property-word.md)** property to apply character and paragraph formatting to the selected text. Use the  **[Font](selection-font-property-word.md)** property to access character formatting properties and methods and the  **[ParagraphFormat](selection-paragraphformat-property-word.md)** property to access paragraph formatting properties and methods.


```vb
Sub FormatSelection() 
 With Selection.Font 
 .Name = "Times New Roman" 
 .Size = 14 
 .AllCaps = True 
 End With 
 With Selection.ParagraphFormat 
 .LeftIndent = InchesToPoints(0.5) 
 .Space1 
 End With 
End Sub
```


## Applying formatting to a range

The following example defines a  **[Range](range-object-word.md)** object that refers to the first three paragraphs in the active document. The  **Range** is formatted by applying properties of the **[Font](font-object-word.md)** object and the  **[ParagraphFormat](paragraphformat-object-word.md)** object.


```vb
Sub FormatRange() 
 Dim rngFormat As Range 
 Set rngFormat = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(1).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(3).Range.End) 
 With rngFormat 
 .Font.Name = "Arial" 
 .ParagraphFormat.Alignment = wdAlignParagraphJustify 
 End With 
End Sub
```


## Inserting text and applying character and paragraph formatting

The following example adds the word "Title" at the top of the current document. The first paragraph is center-aligned and one half-inch space is added after the paragraph. The word "Title" is formatted with 24-point Arial font.


```vb
Sub InsertFormatText() 
 Dim rngFormat As Range 
 Set rngFormat = ActiveDocument.Range(Start:=0, End:=0) 
 With rngFormat 
 .InsertAfter Text:="Title" 
 .InsertParagraphAfter 
 With .Font 
 .Name = "Tahoma" 
 .Size = 24 
 .Bold = True 
 End With 
 End With 
 With ActiveDocument.Paragraphs(1) 
 .Alignment = wdAlignParagraphCenter 
 .SpaceAfter = InchesToPoints(0.5) 
 End With 
End Sub
```


## Switching the space before a paragraph between 12 points and none

The following example toggles the space-before formatting of the first paragraph in the selection. The macro retrieves the current space before value, and if the value is 12 points, the space-before formatting is removed (the  **[SpaceBefore](paragraph-spacebefore-property-word.md)** property is set to zero). If the space-before value is anything other than 12, the  **SpaceBefore** property is set to 12 points.


```vb
Sub ToggleParagraphSpace() 
 With Selection.Paragraphs(1) 
 If .SpaceBefore <> 0 Then 
 .SpaceBefore = 0 
 Else 
 .SpaceBefore = 6 
 End If 
 End With 
End Sub
```


## Switching bold formatting on and off

The following example toggles bold formatting of the selected text.


```vb
Sub ToggleBold() 
 Selection.Font.Bold = wdToggle 
End Sub
```


## Increasing the left margin by 0.5 inch

The following example increases the left and right margins by 0.5 inch. The  **[PageSetup](pagesetup-object-word.md)** object contains all the page setup attributes of a document (such as left margin, bottom margin, and paper size) as properties. The **[LeftMargin](pagesetup-leftmargin-property-word.md)** property is used to return and set the left margin setting. The  **[RightMargin](pagesetup-rightmargin-property-word.md)** property is used to return and set the right margin setting.


```vb
Sub FormatMargins() 
 With ActiveDocument.PageSetup 
 .LeftMargin = .LeftMargin + InchesToPoints(0.5) 
 .RightMargin = .RightMargin + InchesToPoints(0.5) 
 End With 
End Sub
```


