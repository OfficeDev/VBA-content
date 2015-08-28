
# ParagraphFormat.LineRuleWithin Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Determines whether line spacing between base lines is set to a specific number of points or lines. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **LineRuleWithin**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks
<a name="sectionSection1"> </a>

The value of the  **LineRuleWithin** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|Line spacing between base lines is set to a specific number of points.|
| **msoTrue**| Line spacing between base lines is set to a specific number of lines.|

## Example
<a name="sectionSection2"> </a>

This example displays a message box that shows the setting for line spacing for text in shape two on slide one in the active presentation.


```
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

MsgBox "Current line spacing: " &amp; ls &amp; lsUnits
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ParagraphFormat Object](15d495cf-16e2-5cfb-e99c-a551876e3a8a.md)
#### Other resources


 [ParagraphFormat Object Members](c269be7c-ad65-672d-bcac-2a4913346d3e.md)
