
# Range.Style Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the style for the specified object. Read/write  **Variant**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Style**

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


## Remarks
<a name="sectionSection1"> </a>

To set this property, specify the local name of the style, an integer, a  ** [WdBuiltinStyle](9ef433e9-6770-0e20-e1b6-2d9929ffd616.md)** constant, or an object that represents the style. When you return the style for a range that includes more than one style, only the first character or paragraph style is returned.


## Example
<a name="sectionSection2"> </a>

This example displays the style for each character in the selection. 


 **Note**  Each element of the  **Characters** collection is a **Range** object.


```
For each c in Selection.Characters 
 MsgBox c.Style 
Next c
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
