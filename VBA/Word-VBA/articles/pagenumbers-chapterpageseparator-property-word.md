---
title: PageNumbers.ChapterPageSeparator Property (Word)
keywords: vbawd10.chm159776773
f1_keywords:
- vbawd10.chm159776773
ms.prod: word
api_name:
- Word.PageNumbers.ChapterPageSeparator
ms.assetid: f7bd5275-2bb3-fa5f-3a1b-09affa027daf
ms.date: 06/08/2017
---


# PageNumbers.ChapterPageSeparator Property (Word)

Returns or sets the separator character used between the chapter number and the page number. Read/write  **[WdSeparatorType](wdseparatortype-enumeration-word.md)** .


## Syntax

 _expression_ . **ChapterPageSeparator**

 _expression_ An expression that represents a **[PageNumbers](pagenumbers-object-word.md)** object.


## Remarks

Before you can create page numbers that include chapter numbers, the document headings must have a numbered outline format applied that uses styles from the  **Bullets and Numbering** dialog box. To do this in Visual Basic, use the **ApplyListTemplate** method.


## Example

The first part of this example creates a new document, adds chapter titles and page breaks, and then formats the document by using the last numbered outline format listed in the  **Bullets and Numbering** dialog box. The second part of the example adds centered page numbers — including the chapter number — to the header; an en dash separates the chapter number and the page number.


```vb
Dim intLoop As Integer 
Dim hfTemp As HeaderFooter 
 
Documents.Add 
For intLoop = 1 To 5 
 With Selection 
 .TypeParagraph 
 .InsertBreak 
 End With 
Next intLoop 
ActiveDocument.Content.Style = wdStyleHeading1 
ActiveDocument.Content.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(7) 
 
Set hfTemp = ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary) 
With hfTemp.PageNumbers 
 .Add PageNumberAlignment:=wdAlignPageNumberCenter 
 .NumberStyle = wdPageNumberStyleArabic 
 .IncludeChapterNumber = True 
 .HeadingLevelForChapter = 0 
 .ChapterPageSeparator = wdSeparatorEnDash 
End With
```


## See also


#### Concepts


[PageNumbers Collection Object](pagenumbers-object-word.md)

