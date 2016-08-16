
# Paragraphs.Last Property (Word)

Returns a  **Paragraph** object that represents the last item in the collection of paragraphs.


## Syntax

 _expression_ . **Last**

 _expression_ Required. A variable that represents a **[Paragraphs](bdc7a183-2a98-7d47-c86a-5cecd6c91449.md)** collection.


## Example

This example formats the last paragraph in the active document to be right-aligned.


```vb
ActiveDocument.Paragraphs.Last.Alignment = wdAlignParagraphRight
```


## See also


#### Concepts


[Paragraphs Collection Object](bdc7a183-2a98-7d47-c86a-5cecd6c91449.md)
