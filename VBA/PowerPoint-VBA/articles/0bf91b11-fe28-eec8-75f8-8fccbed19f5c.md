
# ParagraphFormat.LineRuleWithin Property (PowerPoint)

Determines whether line spacing between base lines is set to a specific number of points or lines. Read/write.


## Syntax

 _expression_. **LineRuleWithin**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **LineRuleWithin** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Line spacing between base lines is set to a specific number of points.|
|**msoTrue**| Line spacing between base lines is set to a specific number of lines.|

## Example

This example displays a message box that shows the setting for line spacing for text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange.ParagraphFormat

        ls = .SpaceWithin

        If .LineRuleWithin Then

            lsUnits = " lines"

        Else

            lsUnits = " points"

        End If

    End With

End With

MsgBox "Current line spacing: " &; ls &; lsUnits
```


## See also


#### Concepts


[ParagraphFormat Object](15d495cf-16e2-5cfb-e99c-a551876e3a8a.md)
#### Other resources


[ParagraphFormat Object Members](c269be7c-ad65-672d-bcac-2a4913346d3e.md)
