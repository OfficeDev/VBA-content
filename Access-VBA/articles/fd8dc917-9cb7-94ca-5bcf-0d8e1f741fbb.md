
# ComboBox.BackThemeColorIndex Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets or sets a value that represents a color in the applied color theme associated with the  **BackColor** property of the specified object. Read/write **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BackThemeColorIndex**

 _expression_A variable that represents a  ** [ComboBox](1cf508d5-023e-eb38-3991-71e82b2a4e7e.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The  **BackThemeColorIndex** property contains one of the index values listed in the following table.



|**Index Value**|**Description**|
|:-----|:-----|
|0|Text 1|
|1|Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **BackThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## Example
<a name="sectionSection2"> </a>

The following code example sets the Background Color to the Text 2 color by setting the  **BackThemeColorIndex** property.


```
Me.FormHeader.BackThemeColorIndex=2
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ComboBox Object](1cf508d5-023e-eb38-3991-71e82b2a4e7e.md)
#### Other resources


 [ComboBox Object Members](d0d83ca3-3698-295e-5335-7d0816557d6b.md)
