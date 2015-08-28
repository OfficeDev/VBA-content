
# ThreeDFormat.PresetThreeDFormat Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns the preset extrusion format. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **PresetThreeDFormat**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

MsoPresetThreeDFormat


## Remarks
<a name="sectionSection1"> </a>

This property is read-only. To set the preset extrusion format, use the  ** [SetThreeDFormat](9685d3f9-467a-8b11-144a-c4260bdbbddd.md)**method.

Each preset extrusion format contains a set of preset values for the various properties of the extrusion. The values for this property correspond to the options (numbered from left to right, top to bottom) displayed when you click the  **3-D Rotation** submenu on the **Shape Effects** menu.

The value of the  **PresetThreeDFormat** property can be one of these **MsoPresetThreeDFormat** constants. If the value is **msoPresetThreeDFormatMixed**, the extrusion has a custom format rather than a preset format.



| **msoPresetThreeDFormatMixed**|
| **msoThreeD1**|
| **msoThreeD2**|
| **msoThreeD3**|
| **msoThreeD4**|
| **msoThreeD5**|
| **msoThreeD6**|
| **msoThreeD7**|
| **msoThreeD8**|
| **msoThreeD9**|
| **msoThreeD10**|
| **msoThreeD11**|
| **msoThreeD12**|
| **msoThreeD13**|
| **msoThreeD14**|
| **msoThreeD15**|
| **msoThreeD16**|
| **msoThreeD17**|
| **msoThreeD18**|
| **msoThreeD19**|
| **msoThreeD20**|

## Example
<a name="sectionSection2"> </a>

This example sets the extrusion format for shape one on  `myDocument` to 3D Style 12 if the shape initially has a custom extrusion format.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then

        .SetThreeDFormat msoThreeD12

    End If

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ThreeDFormat Object](d6eb7b36-57df-727e-fc5b-50b8c4790c1c.md)
#### Other resources


 [ThreeDFormat Object Members](8d24e2d8-6579-5a14-f403-aaa77b6ed0a6.md)
