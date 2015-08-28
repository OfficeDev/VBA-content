
# CalloutFormat Object (Excel)

 **Last modified:** July 28, 2015

Contains properties and methods that apply to line callouts.

## Remarks

Use the  ** [Callout](80c67ea9-7e55-9841-bbed-302cbd669ce5.md)** property to return a **CalloutFormat** object.


## Example

 The following example specifies the following attributes of shape three (a line callout) on _myDocument_: the callout will have a vertical accent bar that separates the text from the callout line; the angle between the callout line and the side of the callout text box will be 30 degrees; there will be no border around the callout text; the callout line will be attached to the top of the callout text box; and the callout line will contain two segments. For this example to work, shape three must be a callout.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Callout 
 .Accent = True 
 .Angle = msoCalloutAngle30 
 .Border = False 
 .PresetDrop msoCalloutDropTop 
 .Type = msoCalloutThree 
End With
```


## See also


#### Concepts


 [Excel Object Model Reference](11ea8598-8a20-92d5-f98b-0da04263bf2c.md)
#### Other resources


 [CalloutFormat Object Members](29203369-3128-3336-6e78-d1853c4619a6.md)
