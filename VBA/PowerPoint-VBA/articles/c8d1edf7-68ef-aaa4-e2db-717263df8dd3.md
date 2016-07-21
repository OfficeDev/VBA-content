
# TextRange.Copy Method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_ A variable that represents a **TextRange** object.


## Remarks

Use the  **Paste** method to paste the contents of the Clipboard.


## Example

This example copies the text in shape one on slide one in the active presentation to the Clipboard.


```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Copy
```


## See also


#### Concepts


[TextRange Object](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)
#### Other resources


[TextRange Object Members](cb8dc5ff-34de-3d04-1d56-ed387daaf6b9.md)
