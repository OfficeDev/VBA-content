
# PictureFormat Object (Excel)

 **Last modified:** July 28, 2015

Contains properties and methods that apply to pictures and OLE objects.

## Remarks

 The ** [LinkFormat](3d8085bf-c113-7cbe-871b-01f3b6017824.md)** object contains properties and methods that apply to linked OLE objects only. The ** [OLEFormat](96ee06d8-e922-c48c-4406-bb2f5cbaa02a.md)** object contains properties and methods that apply to OLE objects whether or not they're linked.


## Example

Use the  **PictureFormat** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on _myDocument_ and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).PictureFormat 
 .Brightness = 0.3 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayScale 
 .CropBottom = 18
```


## See also


#### Concepts


 [Excel Object Model Reference](11ea8598-8a20-92d5-f98b-0da04263bf2c.md)
#### Other resources


 [PictureFormat Object Members](d27d6074-2698-2b1d-87cb-c9cc187354c3.md)
