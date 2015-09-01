
# Chart Object (Word)

 **Last modified:** July 28, 2015

Represents a chart in a document.

## Remarks

The Example section describes the following properties and methods for returning a  **Chart** object:




- The  ** [Chart](33d577fe-58b0-8e2f-a859-5bd3b34ed892.md)** property.
    
- The  ** [AddChart](http://msdn.microsoft.com/library/1b168e7b-543a-a817-51b0-8171beecc946%28Office.15%29.aspx)** method.
    



## Example

The  ** [InlineShapes](88c632b2-80de-c96a-8879-a98461b38bd0.md)** collection contains an object for each inline shape, including charts, in a document. Use **InlineShapes**( _Index_), where Index is the index number of an inline shape, to return a single  **InlineShape** object. Use the ** [HasChart](f8b88eef-ec41-fc03-f58b-e346d240a121.md)** property to determine whether the **InlineShape** object represents a chart. If the **HasChart** property is set to **True**, use the  ** [Chart](33d577fe-58b0-8e2f-a859-5bd3b34ed892.md)** property to return a **Chart** object.

You can also use the  ** [Type](0f85b99c-025b-9dff-b4f2-b74ab627efcc.md)** property to determine whether the **InlineShape** object represents a chart. If the **Type** property is set to **WdInlineShapeChart**, the inline shape represents a chart.

The following example verifies whether the first inline shape in the active document represents a chart. If so, the example changes the fore color of the first series on the chart.




```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed 
 End If 
End With
```

The following example creates a new 3-D column chart and adds it to the active document.




```
ActiveDocument.InlineShapes.AddChart Type:=xl3DColumn 

```


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [Chart Object Members](8abcbb92-781d-5a42-f395-526cdb3f754e.md)
