
# HeadersFooters.SlideNumber Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns a  ** [HeaderFooter](8aeafb02-adec-17c1-3108-565c78a64ed1.md)** object that represents the slide number in the lower-right corner of a slide, or the page number in the lower-right corner of a notes page or a page of a printed handout or outline. Read-only.

## Syntax

 _expression_. **SlideNumber**

 _expression_A variable that represents a  **HeadersFooters** object.


### Return Value

HeaderFooter


## Example

This example hides the slide number on slide two in the active presentation if the number is currently visible, or it displays the slide number if it is currently hidden.


```
With Application.ActivePresentation.Slides(2) _

        .HeadersFooters.SlideNumber

    If .Visible Then

        .Visible = False

    Else

        .Visible = True

    End If

End With
```


## See also


#### Concepts


 [HeadersFooters Object](5fb10c90-0611-e797-836b-3f18b273af04.md)
#### Other resources


 [HeadersFooters Object Members](b5c50dee-2a19-45fa-0e2b-21620233b5ce.md)
