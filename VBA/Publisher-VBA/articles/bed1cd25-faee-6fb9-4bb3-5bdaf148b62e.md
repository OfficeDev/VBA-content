
# PictureFormat.Brightness Property (Publisher)

Returns or sets a  **Single** indicating the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write.


## Syntax

 _expression_. **Brightness**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

Single


## Remarks

Use the  **[IncrementBrightness](912fd08e-bbb3-bf98-b0da-7128926f3409.md)** method to incrementally adjust the brightness from its current level.


## Example

This example sets the brightness for the first shape in the active publication. The shape must be either a picture or an OLE object.


```vb
ActiveDocument.Pages(1).Shapes(1).PictureFormat _ 
 .Brightness = 0.3
```

