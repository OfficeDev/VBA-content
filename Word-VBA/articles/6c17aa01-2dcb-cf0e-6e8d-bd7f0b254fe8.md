
# Options.InsertedTextMark Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets how Microsoft Word formats inserted text while change tracking is enabled (the  **TrackRevisions** property is **True**). Read/write  **WdInsertedTextMark**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **InsertedTextMark**

 _expression_Required. A variable that represents an  ** [Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** collection.


## Remarks
<a name="sectionSection1"> </a>

If change tracking is not enabled, this property is ignored. Use this property with the  **InsertedTextColor** property to control the appearance of inserted text in a document.

The  **ShowRevisions** property must be **True** to see the formatting for inserted text during editing. The **PrintRevisions** property must be **True** in order for Word to use the formatting for inserted text when printing a document.


## Example
<a name="sectionSection2"> </a>

This example sets Word to italicize inserted text.


```
Options.InsertedTextMark = wdInsertedTextMarkItalic
```

This example sets Word to format inserted text as bold if it isn't already.




```
If Options.InsertedTextMark <> wdInsertedTextMarkBold Then 
 Options.InsertedTextMark = wdInsertedTextMarkBold 
Else 
 MsgBox Prompt:="Inserted text is already bold!" 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


 [Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
