
# Range.Start Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the starting character position of a range. Read/write  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Start**

 _expression_A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


## Remarks
<a name="sectionSection1"> </a>

 **Range** objects have starting and ending character positions. The starting position refers to the character position closest to the beginning of the story. If this property is set to a value larger than that of the **End** property, the **End** property is set to the same value as that of **Start** property.

This property returns the starting character position relative to the beginning of the story. The main text story ( **wdMainTextStory**) begins with character position 0 (zero). You can change the size of a selection, range, or bookmark by setting this property.


## Example
<a name="sectionSection2"> </a>

This example returns the starting position of the second paragraph and the ending position of the fourth paragraph in the active document. The character positions are used to create the range myRange.


```
pos = ActiveDocument.Paragraphs(2).Range.Start 
pos2 = ActiveDocument.Paragraphs(4).Range.End 
Set myRange = ActiveDocument.Range(Start:=pos, End:=pos2)
```

This example moves the starting position of myRange one character to the right (this reduces the size of the range by one character).




```
Set myRange = Selection.Range 
myRange.SetRange Start:=myRange.Start + 1, End:=myRange.End
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
