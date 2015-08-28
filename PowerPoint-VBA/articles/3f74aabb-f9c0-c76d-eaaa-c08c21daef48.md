
# Series.Paste Method (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Paste**

 _expression_A variable that represents a  ** [Series](5c8c2d92-d8ca-4d21-e213-c374292275d4.md)** object.


## Remarks
<a name="sectionSection1"> </a>

You can use this method on column, bar, line, or radar charts, and it sets the  ** [MarkerStyle](e985978e-f0cf-b809-ebe1-f5504e9e8df6.md)** property to **xlMarkerStylePicture**.


## Example
<a name="sectionSection2"> </a>




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example pastes a picture from the Clipboard into series one for the first chart in the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Paste

    End If

End With


```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Series Object](5c8c2d92-d8ca-4d21-e213-c374292275d4.md)
#### Other resources


 [Series Object Members](f7e7168d-3c6f-20db-1e75-56a101c69a70.md)
