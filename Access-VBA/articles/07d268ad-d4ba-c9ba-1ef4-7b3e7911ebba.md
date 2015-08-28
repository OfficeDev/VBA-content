
# Attachment.PictureSizeMode Property (Access)

 **Last modified:** July 28, 2015

You can use the  **PictureSizeMode** property to specify how a picture for an attachment control is sized. Read/write **Byte**.

## Syntax

 _expression_. **PictureSizeMode**

 _expression_A variable that represents an  **Attachment** object.


## Remarks

The  **PictureSizeMode** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Clip|0|(Default) The picture is displayed in its actual size. If the picture is larger than the attachment control, then the picture is clipped.|
|Stretch|1|The picture is stretched horizontally and vertically to fill the entire attachment control, even if its original ratio of height to width is distorted.|
|Zoom|3|The picture is enlarged to the maximum extent possible while keeping its original ratio of height to width.|
When a small picture is used for the  **DefaultPicture**property of an attachment control, setting the  **PictureSizeMode** property to Stretch or Zoom can cause substantial distortion of its resolution. Smaller pictures can be tiled across the entire attachment control by using the **PictureTiling**property.


## See also


#### Concepts


 [Attachment Object](b0756145-9012-f9b9-7df9-e168defed3bf.md)
#### Other resources


 [Attachment Object Members](4294b913-7691-5f45-2c20-5137c2320620.md)
