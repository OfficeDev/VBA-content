
# Trendline.Intercept Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the point where the trendline crosses the value axis. Read/write  **Double**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Intercept**

 _expression_A variable that represents a  ** [Trendline](74755c19-0a7d-cbbf-857e-78740adf6aa4.md)** object.


## Remarks
<a name="sectionSection1"> </a>

Setting this property sets the  ** [InterceptIsAuto](568c57e5-c42f-8559-9c7c-30a72e46463a.md)**property to  **False**.


## Example
<a name="sectionSection2"> </a>




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets trendline one for the first chart in the active document to cross the value axis at 5. You should run the example on a 2-D column chart that contains a single series that has a trendline.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Trendlines(1).Intercept = 5

    End If

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Trendline Object](74755c19-0a7d-cbbf-857e-78740adf6aa4.md)
#### Other resources


 [Trendline Object Members](44b65b82-10a0-ac4b-20f4-890797a20dcf.md)
