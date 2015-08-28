
# Shape.Creator Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Creator**

 _expression_A variable that represents a  **Shape** object.


### Return Value

Long


## Remarks
<a name="sectionSection1"> </a>

The  **Creator** property is designed to be used in Microsoft Office applications for the Macintosh.


## Example
<a name="sectionSection2"> </a>

This example displays a message about the creator of myObject.


```
Set myObject = Application.ActivePresentation.Slides(1).Shapes(1)

If myObject.Creator = &amp;h50575054 Then

    MsgBox "This is a PowerPoint object"

Else

    MsgBox "This is not a PowerPoint object"

End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Shape Object](1da93849-99e0-827e-ced3-c6cf7f8569f3.md)
#### Other resources


 [Shape Object Members](e371c375-c16a-33ef-32b7-6dcb99d3d128.md)
