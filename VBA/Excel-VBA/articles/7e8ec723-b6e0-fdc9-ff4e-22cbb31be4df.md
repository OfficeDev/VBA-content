
# PictureFormat Object (Excel)

Contains properties and methods that apply to pictures and OLE objects.


## Remarks

 The **[LinkFormat](3d8085bf-c113-7cbe-871b-01f3b6017824.md)** object contains properties and methods that apply to linked OLE objects only. The **[OLEFormat](96ee06d8-e922-c48c-4406-bb2f5cbaa02a.md)** object contains properties and methods that apply to OLE objects whether or not they're linked.


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


## Methods



|**Name**|
|:-----|
|[IncrementBrightness](3f75ff17-6cd6-e397-468c-6bf0d1307578.md)|
|[IncrementContrast](6bb72eed-c291-fac2-f4ca-4ca847bd8458.md)|

## Properties



|**Name**|
|:-----|
|[Application](afc9ab72-cf23-a4de-1c21-4d4e28bd623b.md)|
|[Brightness](f17ee171-47da-c982-2f48-9ee333193add.md)|
|[ColorType](6c183163-8fbd-3a0f-b087-05d8d2cdbfd5.md)|
|[Contrast](994cfca5-8ddb-d943-63c8-21abe8508de6.md)|
|[Creator](4a2777a6-ed15-ed24-4553-1b96172ab57f.md)|
|[Crop](229fc83c-488f-887e-5ccf-b900c61ed840.md)|
|[CropBottom](b2c3168f-37db-80a8-815c-b6a2c5a74047.md)|
|[CropLeft](e5d542cb-8653-c798-aede-28c58e4979d6.md)|
|[CropRight](9cf71268-5d63-4f66-6245-968786db14a8.md)|
|[CropTop](adde9cc2-ca09-8494-d250-92a36dfa51e0.md)|
|[Parent](215d013c-02cc-bbe2-32f1-585888506ece.md)|
|[TransparencyColor](c3a7a247-0cc2-adc8-e13f-a1f4ff728ba0.md)|
|[TransparentBackground](9b7cc5b5-610a-821b-cf99-e2af5c4ecf61.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)