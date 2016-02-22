
# Bookmark.Start Property (Word)

Returns or sets the starting character position of a bookmark. Read/write  **Long** .


## Syntax

 _expression_ . **Start**

 _expression_ A variable that represents a **[Bookmark](be6b0c7b-60ca-97e7-ef19-6de335da3197.md)** object.


## Remarks

If this property is set to a value larger than that of the  **[End](05531b0d-b05e-0010-9ff8-ba6d90de560d.md)** property, the **End** property is set to the same value as that of **Start** property.

 Bookmark objects have starting and ending character positions. The starting position refers to the character position closest to the beginning of the story.

This property returns the starting character position relative to the beginning of the story. The main text story ( **wdMainTextStory** ) begins with character position 0 (zero). You can change the size of a bookmark by setting this property.


## Example

This example compares the ending position of the "temp" bookmark with the starting position of the "begin" bookmark.


```vb
Set Book1 = ActiveDocument.Bookmarks("begin") 
Set Book2 = ActiveDocument.Bookmarks("temp") 
If Book2.End > Book1.Start Then Book1.Select
```


## See also


#### Concepts


[Bookmark Object](be6b0c7b-60ca-97e7-ef19-6de335da3197.md)
#### Other resources


[Bookmark Object Members](c7ff0d52-501c-64ac-0034-b0e4ed3640f2.md)
