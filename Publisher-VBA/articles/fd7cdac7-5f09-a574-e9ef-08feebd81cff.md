
# CalloutFormat.Gap Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets a  **Variant** indicating the horizontal distance between the end of the callout line and the text bounding box. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Gap**

 _expression_A variable that represents a  **CalloutFormat** object.


### Return Value

Variant


## Remarks
<a name="sectionSection1"> </a>

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example
<a name="sectionSection2"> </a>

This example sets the distance between the callout line and the text bounding box to 3 points for the first shape in the active publication. For the example to work, the shape must be a callout.


```
ActiveDocument.Pages(1).Shapes(1).Callout.Gap = 3
```

