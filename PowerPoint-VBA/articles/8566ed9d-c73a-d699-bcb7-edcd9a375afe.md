
# TextRange.TrimText Method (PowerPoint)

 **Last modified:** July 28, 2015

Returns a  **TextRange** object that represents the specified text minus any trailing spaces.

## Syntax

 _expression_. **TrimText**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

TextRange


## Example

This example inserts the string " Text to trim " at the beginning of the text in shape two on slide one in the active presentation and then displays message boxes showing the string before and after it is trimmed.


```
With Application.ActivePresentation.Slides(1).Shapes(2) _

        .TextFrame.TextRange

    With .InsertBefore("   Text to trim   ")

        MsgBox "Untrimmed: " &amp; """" &amp; .Text &amp; """"

        MsgBox "Trimmed: " &amp; """" &amp; .TrimText.Text &amp; """"

    End With

End With
```


## See also


#### Concepts


 [TextRange Object](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)
#### Other resources


 [TextRange Object Members](cb8dc5ff-34de-3d04-1d56-ed387daaf6b9.md)
