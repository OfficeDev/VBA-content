
# TextRange.InsertBefore Method (PowerPoint)

 **Last modified:** July 28, 2015

Appends a string to the beginning of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.

## Syntax

 _expression_. **InsertBefore**( **_NewText_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|NewText|Optional| **String**|The text to be appended. The default value is an empty string.|

## Example

This example appends the string "Test version: " to the beginning of the title on slide one in the active presentation.


```
With Application.ActivePresentation.Slides(1).Shapes(1)

    .TextFrame.TextRange.InsertBefore "Test version: "

End With
```

This example appends the contents of the Clipboard to the beginning of the title on slide one in the active presentation.




```
Application.ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange.InsertBefore.Paste
```


## See also


#### Concepts


 [TextRange Object](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)
#### Other resources


 [TextRange Object Members](cb8dc5ff-34de-3d04-1d56-ed387daaf6b9.md)
