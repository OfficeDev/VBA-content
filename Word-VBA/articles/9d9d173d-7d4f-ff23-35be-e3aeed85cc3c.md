
# Paragraphs.Last Property (Word)

 **Last modified:** July 28, 2015

Returns a  **Paragraph** object that represents the last item in the collection of paragraphs.

## Syntax

 _expression_. **Last**

 _expression_Required. A variable that represents a  ** [Paragraphs](bdc7a183-2a98-7d47-c86a-5cecd6c91449.md)** collection.


## Example

This example formats the last paragraph in the active document to be right-aligned.


```
ActiveDocument.Paragraphs.Last.Alignment = wdAlignParagraphRight
```


## See also


#### Concepts


 [Paragraphs Collection Object](bdc7a183-2a98-7d47-c86a-5cecd6c91449.md)
#### Other resources


 [Paragraphs Object Members](490e2695-3cdd-4906-f730-583d18486aa2.md)
