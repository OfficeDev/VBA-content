
# Range.TextVisibleOnScreen Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Property value](#sectionSection2)


Returns a  **Long** that indicates whether the text in the specified range is visible on the screen. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **TextVisibleOnScreen**

 _expression_A variable that represents a  **Range** object.


## Remarks
<a name="sectionSection1"> </a>

The  **TextVisibleOnScreen** property returns 1 if all text in the range is visible; it returns 0 if no text in the range is visible; and it returns -1 if some text in the range is visible and some is not. Text that is not visible could be, for example, text that is in a collapsed heading.


## Property value
<a name="sectionSection2"> </a>

 **INT32**


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
