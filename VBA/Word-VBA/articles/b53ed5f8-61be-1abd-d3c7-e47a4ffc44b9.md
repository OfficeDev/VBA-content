
# FillFormat.PresetGradientType Property (Word)

Returns the preset gradient type for the specified fill. Read-only  **MsoPresetGradientType** .


## Syntax

 _expression_ . **PresetGradientType**

 _expression_ An expression that represents a **[FillFormat](39205d07-9e37-1be1-ec4a-93ba8bac2f26.md)** object.


## Remarks

Use the  **[PresetGradient](bffe754d-6593-9684-abf4-b5d1e9df720e.md)** method to set the preset gradient type for the fill.


## Example

This example changes the fill for all shapes in  `myDocument` with the Moss preset gradient fill to the Fog preset gradient fill.


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 With s.Fill 
 If .PresetGradientType = msoGradientMoss Then 
 .PresetGradient msoGradientHorizontal, 1, _ 
 msoGradientFog 
 End If 
 End With 
Next
```


## See also


#### Concepts


[FillFormat Object](39205d07-9e37-1be1-ec4a-93ba8bac2f26.md)
#### Other resources


[FillFormat Object Members](09251952-b63e-4886-d2fa-938e27dba15a.md)
