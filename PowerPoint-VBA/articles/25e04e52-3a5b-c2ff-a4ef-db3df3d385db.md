
# ShapeNode.Creator Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Creator**

 _expression_A variable that represents a  **ShapeNode** object.


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


 [ShapeNode Object](031edfef-4eae-39b2-0c73-90d2065741aa.md)
#### Other resources


 [ShapeNode Object Members](b9840b71-bba6-e7b0-c4c4-943bd306d9bd.md)
