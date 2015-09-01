
# Series.Paste Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Paste**

 _expression_A variable that represents a  ** [Series](212c323f-8acb-2ba7-1359-ab0f43268e77.md)** object.


## Remarks
<a name="sectionSection1"> </a>

You can use this method on column, bar, line, or radar charts, and it sets the  ** [MarkerStyle](d9ba7847-2785-0f29-7e6e-d4bb2d62fc2f.md)** property to **xlMarkerStylePicture**.


## Example
<a name="sectionSection2"> </a>

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


 [Series Object](212c323f-8acb-2ba7-1359-ab0f43268e77.md)
#### Other resources


 [Series Object Members](0bc84851-3f0a-15e0-ae2b-c36215709220.md)
