
# PageNumbers.ChapterPageSeparator Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the separator character used between the chapter number and the page number. Read/write  ** [WdSeparatorType](94cb01b0-b850-ddc6-46ce-ea0261d38247.md)**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ChapterPageSeparator**

 _expression_An expression that represents a  ** [PageNumbers](9090f96e-d898-ace6-35fa-f6e59c527ea2.md)** object.


## Remarks
<a name="sectionSection1"> </a>

Before you can create page numbers that include chapter numbers, the document headings must have a numbered outline format applied that uses styles from the  **Bullets and Numbering** dialog box. To do this in Visual Basic, use the **ApplyListTemplate** method.


## Example
<a name="sectionSection2"> </a>

The first part of this example creates a new document, adds chapter titles and page breaks, and then formats the document by using the last numbered outline format listed in the  **Bullets and Numbering** dialog box. The second part of the example adds centered page numbers â€” including the chapter number â€” to the header; an en dash separates the chapter number and the page number.


```
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
<a name="sectionSection2"> </a>


#### Concepts


 [PageNumbers Collection Object](9090f96e-d898-ace6-35fa-f6e59c527ea2.md)
#### Other resources


 [PageNumbers Object Members](7f6d35df-499d-b3bf-6eaa-70e2ab1a2e8d.md)
