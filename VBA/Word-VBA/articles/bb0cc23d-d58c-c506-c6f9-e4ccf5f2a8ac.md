
# Options.PictureWrapType Property (Word)

Sets or returns a  **WdWrapTypeMerged** that indicates how Microsoft Word wraps text around pictures. Read/write.


## Syntax

 _expression_ . **PictureWrapType**

 _expression_ Required. A variable that represents an **[Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** collection.


## Remarks

This is a default option setting and affects all pictures inserted unless picture wrapping is individually defined for a picture.


## Example

This example sets Word to insert and paste all pictures inline with the text if inline is not already specified.


```vb
Sub PicWrap() 
 With Application.Options 
 If .PictureWrapType <> wdWrapMergeInline Then 
 .PictureWrapType = wdWrapMergeInline 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


[Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
