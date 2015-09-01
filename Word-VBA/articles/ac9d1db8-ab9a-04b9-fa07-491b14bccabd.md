
# Border.Color Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the 24-bit color for the specified  **Border** object.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Color**

 _expression_Required. A variable that represents a  ** [Border](be48c020-b86c-c004-ce1c-76d9edae9791.md)** object.


## Remarks
<a name="sectionSection1"> </a>

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.


## Example
<a name="sectionSection2"> </a>

This example adds a dotted indigo border around each cell in the first table.


```
If ActiveDocument.Tables.Count >= 1 Then 
 For Each aBorder In ActiveDocument.Tables(1).Borders 
 aBorder.Color = wdColorIndigo 
 aBorder.LineStyle = wdLineStyleDashDot 
 aBorder.LineWidth = wdLineWidth075pt 
 Next aBorder 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Border Object](be48c020-b86c-c004-ce1c-76d9edae9791.md)
#### Other resources


 [Border Object Members](0c2f320b-8f74-961b-297e-dc068db579aa.md)
