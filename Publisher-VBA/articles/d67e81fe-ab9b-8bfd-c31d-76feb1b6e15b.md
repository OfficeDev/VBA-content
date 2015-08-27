
# TextRange.MajorityParagraphFormat Property (Publisher)

 **Last modified:** July 28, 2015

Returns a  ** [ParagraphFormat](0e5b1c20-564e-ef5c-f24d-1143dcaadcd8.md)** object that represents the paragraph formatting applied to most of the paragraphs in a text range.

## Syntax

 _expression_. **MajorityParagraphFormat**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

ParagraphFormat


## Example

This example applies the paragraph formatting applied to a majority of the paragraphs in the first shape to the paragraphs in the second shape on the first page of the active document. This example assumes that there are at least two shapes on page one of the active publication.


```
Sub SetFontName() 
 Dim fmt As ParagraphFormat 
 With ActiveDocument.Pages(1) 
 Set fmt = .Shapes(1).TextFrame.TextRange _ 
 .MajorityParagraphFormat 
 .Shapes(2).TextFrame.TextRange.ParagraphFormat = fmt 
 End With 
End Sub
```

