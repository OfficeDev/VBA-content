
# CalloutFormat.Drop Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only  **Single**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Drop**

 _expression_A variable that represents a  **CalloutFormat** object.


## Remarks
<a name="sectionSection1"> </a>

This distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.

Use the  ** [CustomDrop](d38513f6-1c42-e4b3-7a0f-b8543d59d0ff.md)**method to set the value of this property.

The value of this property accurately reflects the position of the callout line attachment to the text box only if the callout has an explicitly set drop value â€” that is, if the value of the  ** [DropType](ab947fa4-4af9-e491-f62d-e0ca036e1892.md)**property is  **msoCalloutDropCustom**.


## Example
<a name="sectionSection2"> </a>

This example replaces the custom drop for shape one on  `myDocument` with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box. For the example to work, shape one must be a callout.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Callout 
 If .DropType = msoCalloutDropCustom Then 
 If .Drop < .Parent.Height / 2 Then 
 .PresetDrop msoCalloutDropTop 
 Else 
 .PresetDrop msoCalloutDropBottom 
 End If 
 End If 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [CalloutFormat Object](d9d7d279-04ef-dbee-23cd-ddd606ed917d.md)
#### Other resources


 [CalloutFormat Object Members](29203369-3128-3336-6e78-d1853c4619a6.md)
