
# Font.NameComplexScript Property (PowerPoint)

Returns or sets the complex script font name. Used for mixed language text. Read/write.


## Syntax

 _expression_. **NameComplexScript**

 _expression_ A variable that represents a **Font** object.


### Return Value

String


## Remarks

When you have a right-to-left language setting specified, this property is equivalent to the  **Complex scripts font** list in the **Font** dialog box ( **Font** tab). When you have an Asian language setting specified, this property is equivalent to the **Asian text font** list in the **Font** dialog box ( **Font** tab).


## Example

This example sets the complex script font to Times New Roman.


```vb
ActivePresentation.Slides(1).Shapes.Title.TextFrame _
    .TextRange.Font.NameComplexScript = "Times New Roman"
```


## See also


#### Concepts


[Font Object](ad62daaa-01a5-36cc-5451-e0da0134ac95.md)
#### Other resources


[Font Object Members](a2043117-2222-dad3-d73c-0e9d5591c9be.md)
