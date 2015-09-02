
# DefaultWebOptions Object (Excel)

 **Last modified:** July 28, 2015

Contains global application-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page. You can return or set attributes either at the application (global) level or at the workbook level.

## Remarks

 Workbook-level attribute settings override application-level attribute settings. Workbook-level attributes are contained in the ** [WebOptions](d573637f-1891-4602-c961-091795e47356.md)** object.


 **Note**  Attribute values can be different from one workbook to another, depending on the attribute value at the time the workbook was saved.


## Example

Use the  ** [DefaultWebOptions](51524888-0812-85ee-c8f9-e14d9b558f57.md)** property to return the **DefaultWebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and sets the _strImageFileType_ variable accordingly.


```
Set objAppWebOptions = Application.DefaultWebOptions 
With objAppWebOptions 
 If .AllowPNG = True Then 
 strImageFileType = "PNG" 
 Else 
 strImageFileType = "JPG" 
 End If 
End With
```


## See also


#### Concepts


 [Excel Object Model Reference](11ea8598-8a20-92d5-f98b-0da04263bf2c.md)
#### Other resources


 [DefaultWebOptions Object Members](52db1398-01d8-eba5-772f-2923fdc89f5b.md)
