
# Paragraph.AddSpaceBetweenFarEastAndAlpha Property (Word)

 **Last modified:** July 28, 2015

 **True** if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long**.

## Syntax

 _expression_. **AddSpaceBetweenFarEastAndAlpha**

 _expression_A variable that represents a  ** [Paragraph](0a704079-a082-4ab1-841b-fc0d49dd26d4.md)** object.


## Example

This example sets Microsoft Word to automatically add spaces between Japanese and Latin text for the first paragraph in the active document.


```
ActiveDocument.Paragraphs(1).AddSpaceBetweenFarEastAndAlpha = True
```


## See also


#### Concepts


 [Paragraph Object](0a704079-a082-4ab1-841b-fc0d49dd26d4.md)
#### Other resources


 [Paragraph Object Members](e1fc5b91-e908-580e-ab72-898648a5c0c3.md)
