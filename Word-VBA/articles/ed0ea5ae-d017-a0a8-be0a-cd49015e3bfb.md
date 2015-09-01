
# ListLevel.TextPosition Property (Word)

 **Last modified:** July 28, 2015

Returns or sets the position (in points) for the second line of wrapping text for the specified  **ListLevel**object. Read/write  **Single**.

## Syntax

 _expression_. **TextPosition**

 _expression_An expression that returns a  ** [ListLevel](0cd152cb-6c25-50cb-7c1d-8b6d9734505b.md)** object.


## Example

This example sets the indentation for all levels of the first outline-numbered list template. Each list level number is indented 0.5 inch (36 points) from the previous level, the tab is set at 0.25 inch (18 points) from the number, and wrapping text is indented 0.25 inch (18 points) from the number.


```
r = 0 
For Each lev In ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(1).ListLevels 
 lev.Alignment = wdListLevelAlignLeft 
 lev.NumberPosition = r 
 lev.TrailingCharacter = wdTrailingTab 
 lev.TabPosition = r + 18 
 lev.TextPosition = r + 18 
 r = r + 36 
Next lev
```


## See also


#### Concepts


 [ListLevel Object](0cd152cb-6c25-50cb-7c1d-8b6d9734505b.md)
#### Other resources


 [ListLevel Object Members](befd48fb-74b1-e505-a027-af8534e02f19.md)
