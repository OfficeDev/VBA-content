
# ParagraphFormat.HangingPunctuation Property (PowerPoint)

Returns or sets the hanging punctuation option if you have an Asian language setting specified. Read/write.


## Syntax

 _expression_. **HangingPunctuation**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HangingPunctuation** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The hanging punctuation option is not selected.|
|**msoTrue**| The hanging punctuation option is selected.|

## Example

This example selects hanging punctuation for the first paragraph of the active presentation.


```vb
ActivePresentation.Paragraphs(1).HangingPunctuation = msoTrue
```


## See also


#### Concepts


[ParagraphFormat Object](15d495cf-16e2-5cfb-e99c-a551876e3a8a.md)
#### Other resources


[ParagraphFormat Object Members](c269be7c-ad65-672d-bcac-2a4913346d3e.md)
