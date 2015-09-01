
# PageNumbers.IncludeChapterNumber Property (Word)

 **Last modified:** July 28, 2015

 **True** if a chapter number is included with page numbers or a caption label. Read/write **Boolean**.

## Syntax

 _expression_. **IncludeChapterNumber**

 _expression_A variable that represents a  ** [PageNumbers](9090f96e-d898-ace6-35fa-f6e59c527ea2.md)** object.


## Example

This example adds page numbers in the footer for section one in the active document. The page numbers include the chapter number.


```
With ActiveDocument.Sections(1).Footers _ 
 (wdHeaderFooterPrimary).PageNumbers 
 .Add 
 .IncludeChapterNumber = True 
 .HeadingLevelForChapter = 1 
End With
```

This example adds the chapter number from the Heading 2 style to figure captions, sets the caption numbering style, and then inserts a new figure caption. The document should already contain a Heading 2 style with numbering.




```
With CaptionLabels(wdCaptionFigure) 
 .IncludeChapterNumber = True 
 .ChapterStyleLevel = 2 
 .NumberStyle = wdCaptionNumberStyleUppercaseLetter 
End With 
Selection.InsertCaption Label:="Figure", Title:=": History"
```


## See also


#### Concepts


 [PageNumbers Collection Object](9090f96e-d898-ace6-35fa-f6e59c527ea2.md)
#### Other resources


 [PageNumbers Object Members](7f6d35df-499d-b3bf-6eaa-70e2ab1a2e8d.md)
