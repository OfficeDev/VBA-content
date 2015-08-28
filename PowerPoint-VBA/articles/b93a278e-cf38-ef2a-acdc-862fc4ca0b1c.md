
# ChartFont.FontStyle Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the font style. Read/write  **String**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **FontStyle**

 _expression_A variable that represents a  ** [ChartFont](185dfaa0-4ed9-01d2-6584-b0838b50ef8c.md)** object.


## Remarks
<a name="sectionSection1"> </a>

Changing this property may affect other  **ChartFont** properties (such as ** [Bold](5d5a0b2e-5aab-f197-79da-e9bb8d219af9.md)** and ** [Italic](c62ad4c5-c7b3-58d8-8d37-540a8a123ce2.md)**).


## Example
<a name="sectionSection2"> </a>




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font style for the title of the first chart in the active document to bold and italic.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Title.Font.FontStyle = "Bold Italic"

    End If

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ChartFont Object](185dfaa0-4ed9-01d2-6584-b0838b50ef8c.md)
#### Other resources


 [ChartFont Object Members](8ec251bd-d4f8-bd15-0b7f-5da95409d315.md)
