
# Legend.LegendEntries Method (PowerPoint)

 **Last modified:** July 28, 2015

Returns a collection of legend entries for the legend.

## Syntax

 _expression_. **LegendEntries**

 _expression_A variable that represents a  ** [Legend](7be25694-8694-049a-c31f-533fe6fd0562.md)** object.


### Return Value

A  ** [LegendEntries](ac65aeaa-8a1c-57d7-499f-1c0b57dd02fd.md)** object that represents the legend entries for the legend.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font for legend entry one on the first chart in the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Legend.LegendEntries(1).Font.Name = "Arial"

    End If

End With
```


## See also


#### Concepts


 [Legend Object](7be25694-8694-049a-c31f-533fe6fd0562.md)
#### Other resources


 [Legend Object Members](138eddc7-3b48-bc0a-163b-3e6f7560ed97.md)
