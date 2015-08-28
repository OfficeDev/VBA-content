
# ComboBox.ReadingOrder Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


You can use the  **ReadingOrder** property to specify or determine the reading order of words in text. Read/write **Byte**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ReadingOrder**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks
<a name="sectionSection1"> </a>

The  **ReadingOrder** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Context|0|Reading order is determined by the language of the first character entered. If a right-to-left language character is entered first, reading order is right to left. If a left-to-right language character is entered first, reading order is left to right.|
|Left-to-Right|1|Sets the reading order to left to right.|
|Right-to-Left|2|Sets the reading order to right to left.|
In a combo box or list box, the  **ReadingOrder** property determines reading order behavior for both the text box and list box components of the control.


## Example
<a name="sectionSection2"> </a>

The following example sets the reading order to right to left for the "Address" text box on the "International Shipping" form.


```
Forms("International Shipping").Controls("Address").ReadingOrder = 2
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ComboBox Object](1cf508d5-023e-eb38-3991-71e82b2a4e7e.md)
#### Other resources


 [ComboBox Object Members](d0d83ca3-3698-295e-5335-7d0816557d6b.md)
