
# HeaderFooter.LinkToPrevious Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if the specified header or footer is linked to the corresponding header or footer in the previous section. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **LinkToPrevious**

 _expression_An expression that returns a  ** [HeaderFooter](3f2f926a-9220-5536-80ed-af63d2feb016.md)**object.


## Remarks
<a name="sectionSection1"> </a>

When a header or footer is linked, its contents are the same as in the previous header or footer. Because the  **LinkToPrevious** property is set to **True** by default, you can add headers, footers, and page numbers to your entire document by working with the headers, footers, and page numbers in the first section. For instance, the following example adds page numbers to the header on all pages in all sections of the active document.


```
ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary).PageNumbers.Add
```

The  **LinkToPrevious** property applies to each header or footer individually. For example, the **LinkToPrevious** property could be set to **True** for the even-numbered-page header but **False** for the even-numbered-page footer.


## Example
<a name="sectionSection2"> </a>

The first part of this example creates a new document with two sections. The second part creates unique headers for even-numbered and odd-numbered pages in sections one and two in the new document.


```
Documents.Add 
With Selection 
 For j = 1 to 4 
 .TypeParagraph 
 .InsertBreak 
 .TypeParagraph 
 Next j 
End With 
With ActiveDocument 
 .Paragraphs(5).Range.InsertBreak Type:=wdSectionBreakNextPage 
 .PageSetup.OddAndEvenPagesHeaderFooter = True 
End With 
With ActiveDocument.Sections(2) 
 With .Headers(wdHeaderFooterPrimary) 
 .LinkToPrevious = False 
 .Range.InsertBefore "Section 2 Odd Header" 
 End With 
 With .Headers(wdHeaderFooterEvenPages) 
 .LinkToPrevious = False 
 .Range.InsertBefore "Section 2 Even Header" 
 End With 
End With 
With ActiveDocument.Sections(1) 
 .Headers(wdHeaderFooterPrimary) _ 
 .Range.InsertBefore "Section 1 Odd Header" 
 .Headers(wdHeaderFooterEvenPages) _ 
 .Range.InsertBefore "Section 1 Even Header" 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [HeaderFooter Object](3f2f926a-9220-5536-80ed-af63d2feb016.md)
#### Other resources


 [HeaderFooter Object Members](400647fc-cf49-a468-850f-f94a054552c0.md)
