
# Shape.ScaleWidth Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Scales the width of the shape by a specified factor.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ScaleWidth**( **_Factor_**,  **_RelativeToOriginalSize_**,  **_Scale_**)

 _expression_Required. A variable that represents a  ** [Shape](604029ce-9b2f-9748-5d4e-b458796fa2f0.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Factor|Required| **Single**|Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
|RelativeToOriginalSize|Required| **MsoTriState**| **True** to scale the shape relative to its original size. **False** to scale it relative to its current size. You can specify **True** for this argument only if the specified shape is a picture or an OLE object.|
|Scale|Optional| **MsoScaleFrom**|The part of the shape that retains its position when the shape is scaled.|

## Remarks
<a name="sectionSection1"> </a>

For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.


## Example
<a name="sectionSection2"> </a>

This example scales all pictures and OLE objects on  _myDocument_ to 175 percent of their original height and width, and it scales all other shapes to 175 percent of their current height and width.


```
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 Select Case s.Type 
 Case msoEmbeddedOLEObject, msoLinkedOLEObject, _ 
 msoOLEControlObject, _ 
 msoLinkedPicture, msoPicture 
 s.ScaleHeight 1.75, True 
 s.ScaleWidth 1.75, True 
 Case Else 
 s.ScaleHeight 1.75, False 
 s.ScaleWidth 1.75, False 
 End Select 
Next
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Shape Object](604029ce-9b2f-9748-5d4e-b458796fa2f0.md)
#### Other resources


 [Shape Object Members](4aa8e2f4-5629-3922-11e4-df028bd1e1de.md)
