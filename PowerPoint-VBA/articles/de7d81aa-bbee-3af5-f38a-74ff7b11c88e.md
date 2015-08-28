
# Point.Explosion Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Read/write  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Explosion**

 _expression_A variable that represents a  ** [Point](e0137fdd-5632-88d7-a6c0-57a76717e736.md)** object.


## Remarks
<a name="sectionSection1"> </a>

This property returns 0 (zero) if there is no explosion (the tip of the slice is in the center of the pie). 


## Example
<a name="sectionSection2"> </a>




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the explosion value for point two of the first chart in the active document. You should run the example on a pie chart.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Points(2).Explosion = 20

    End If

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Point Object](e0137fdd-5632-88d7-a6c0-57a76717e736.md)
#### Other resources


 [Point Object Members](ddf0303f-d97f-91fd-12b5-e569a7899ebd.md)
