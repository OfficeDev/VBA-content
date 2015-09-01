
# Page.PictureTiling Property (Outlook Forms Script)

 **Last modified:** July 28, 2015

Returns or sets a  **Boolean** that specifies whether a picture is repeated across the background of the object. Read/write.

## Syntax

 _expression_. **PictureTiling**

 _expression_A variable that represents a  **Page** object.


## Remarks

 **True** if the picture is tiled across the background, **False** otherwise (default).

If a picture is smaller than the form or page that contains it, you can tile the picture on the form or page.

The tiling pattern depends on the current setting of the  ** [PictureAlignment](c52f0b5b-c703-d9d6-1bae-e4fe9b696cf8.md)** and ** [PictureSizeMode](24a0415a-f89a-c0fb-9c44-b33484c8cd49.md)** properties. For example, if **PictureAlignment** is set to 0, the tiling pattern starts at the upper left and repeats the picture across the form or page and down the height of the form or page. If **PictureSizeMode** is set to 0, the tiling pattern crops the last tile if it doesn't completely fit on the form or page.

