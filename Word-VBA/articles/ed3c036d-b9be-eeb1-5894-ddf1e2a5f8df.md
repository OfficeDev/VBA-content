
# ListTemplate.ListLevels Property (Word)

Returns a  **[ListLevels](9165c008-c066-8d3e-9254-d9e0ab2ec091.md)** collection that represents all the levels for the specified **ListTemplate** .


## Syntax

 _expression_ . **ListLevels**

 _expression_ An expression that returns a **[ListTemplate](d5e339f7-5798-305b-a6b0-6b572d9112f4.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the variable myListTemp to the first list template (excluding None) on the  **Outline Numbered** tab in the **Bullets and Numbering** dialog box ( **Format** menu). Each level in the list has a matching heading style linked to it.


```vb
Set myListTemp = _ 
 ListGalleries(wdOutlineNumberGallery).ListTemplates(1) 
For Each mylevel In myListTemp.ListLevels 
 mylevel.LinkedStyle = "Heading " &; mylevel.index 
Next mylevel
```


## See also


#### Concepts


[ListTemplate Object](d5e339f7-5798-305b-a6b0-6b572d9112f4.md)
#### Other resources


[ListTemplate Object Members](d084eb01-aeeb-259b-91c5-5268fe0395c9.md)
