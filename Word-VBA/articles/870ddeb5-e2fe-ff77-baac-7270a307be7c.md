
# Paragraph.Shading Property (Word)

 **Last modified:** July 28, 2015

Returns a  ** [Shading](e136509a-1be1-29e4-7b37-1faf659e37ba.md)** object that refers to the shading formatting for the specified paragraph.

## Syntax

 _expression_. **Shading**

 _expression_Required. A variable that represents a  ** [Paragraph](0a704079-a082-4ab1-841b-fc0d49dd26d4.md)** object.


## Example

This example applies yellow shading to the first paragraph in the selection.


```
With Selection.Paragraphs(1).Shading 
 .Texture = wdTexture12Pt5Percent 
 .BackgroundPatternColorIndex = wdYellow 
 .ForegroundPatternColorIndex = wdBlack 
End With
```


## See also


#### Concepts


 [Paragraph Object](0a704079-a082-4ab1-841b-fc0d49dd26d4.md)
#### Other resources


 [Paragraph Object Members](e1fc5b91-e908-580e-ab72-898648a5c0c3.md)
