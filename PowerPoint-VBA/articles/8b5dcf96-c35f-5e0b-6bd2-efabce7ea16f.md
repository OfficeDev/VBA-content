
# ParagraphFormat.SpaceAfter Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns or sets the amount of space after the last line in each paragraph of the specified text, in points or lines. Read/write.

## Syntax

 _expression_. **SpaceAfter**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Single


## Example

This example sets the spacing after paragraphs to 6 points for the text in shape two on slide one in the active presentation.


```
With Application.ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat

        .LineRuleAfter = False

        .SpaceAfter = 6

    End With

End With
```


## See also


#### Concepts


 [ParagraphFormat Object](15d495cf-16e2-5cfb-e99c-a551876e3a8a.md)
#### Other resources


 [ParagraphFormat Object Members](c269be7c-ad65-672d-bcac-2a4913346d3e.md)
