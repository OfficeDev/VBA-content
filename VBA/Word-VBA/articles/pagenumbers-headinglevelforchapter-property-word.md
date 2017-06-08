---
title: PageNumbers.HeadingLevelForChapter Property (Word)
keywords: vbawd10.chm159776772
f1_keywords:
- vbawd10.chm159776772
ms.prod: word
api_name:
- Word.PageNumbers.HeadingLevelForChapter
ms.assetid: 2202f3de-9ce0-e8d9-ad7c-9c95c1bc8f26
ms.date: 06/08/2017
---


# PageNumbers.HeadingLevelForChapter Property (Word)

Returns or sets the heading level style that's applied to the chapter titles in the document. Read/write  **Long** .


## Syntax

 _expression_ . **HeadingLevelForChapter**

 _expression_ A variable that represents a **[PageNumbers](pagenumbers-object-word.md)** object.


## Remarks

The  **HeadingLevelForChapter** property can be a number from 0 (zero) through 8, corresponding to heading levels 1 through 9.

Before you can create page numbers that include chapter numbers, the document headings must have a numbered outline format applied that uses styles from the  **Bullets and Numbering** dialog box. To do this in Visual Basic, use the **ApplyListTemplate** method.


## Example

The first part of this example creates a new document, adds chapter titles and page breaks, and then formats the document by using the last numbered outline format listed in the  **Bullets and Numbering** dialog box. The second part of the example adds centered page numbers - including the chapter number - to the header; an en dash separates the chapter number and the page number. The first heading level is used for the chapter number, and lowercase roman numerals are used for the page number.


```vb
Dim intLoop As Integer 
Dim hdrTemp As HeaderFooter 
 
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
 
Set hdrTemp = ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary) 
 
With hdrTemp.PageNumbers 
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

