
# ColorsInUse Object (Publisher)

 **Last modified:** July 28, 2015

A collection of  ** [ColorFormat](659069e1-e359-94d7-de06-a1d98378193b.md)** objects that represent the colors present in the specified publication.

## Remarks

The  **ColorsInUse** collection supports all the publication color models: RGB, process colors, and spot color.

For process color and spot color publications, colors are based on inks. For a given ink, a publication may contain several colors that are different tints or shades of that ink. Use the  ** [Plates](7da44b06-c94f-dadc-da91-09b757d5a076.md)** collection to access the plates that represent the inks defined for a publication.


## Example

Use the  ** [ColorsInUse](http://msdn.microsoft.com/library/b018ffbc-b848-c0d0-19fa-df053e45260d%28Office.15%29.aspx)** property of the ** [Document](44f02255-ff5b-bcfe-900f-61c8fdf61ef3.md)** object to return the **ColorsInUse** collection.

The following example lists properties of each color in the active publication that is based on the specified ink. This example assumes the publication's color mode has been defined as spot color or process and spot color.




```
Sub ListColorsBasedOnInk() 
Dim cfLoop As ColorFormat 
 
For Each cfLoop In ActiveDocument.ColorsInUse 
 
 With cfLoop 
 If .Ink = "2" Then 
 Debug.Print "BaseRGB: " &amp; .BaseRGB 
 Debug.Print "RGB: " &amp; .RGB 
 Debug.Print "TintShade: " &amp; .TintAndShade 
 Debug.Print "Type: " &amp; .Type 
 End If 
 End With 
 
Next cfLoop 
 
End Sub
```

Use  **ColorsInUse**(index), where index is the color index number, to return a single  **ColorFormat** object. The following example returns properties for the second color in the publication.




```
Sub ColorProperties() 
 
 With ActiveDocument.ColorsInUse(2) 
 Debug.Print "Color RBG: " &amp; .RGB 
 Debug.Print "Ink RBG: " &amp; .BaseRGB 
 Debug.Print "Tint: " &amp; .TintAndShade 
 
 End With 
 
End Sub
```

