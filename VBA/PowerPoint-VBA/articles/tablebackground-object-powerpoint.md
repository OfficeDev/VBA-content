---
title: TableBackground Object (PowerPoint)
keywords: vbapp10.chm677000
f1_keywords:
- vbapp10.chm677000
ms.prod: powerpoint
api_name:
- PowerPoint.TableBackground
ms.assetid: ba29d6df-f37c-05c1-4e29-8c1766a8aaf4
ms.date: 06/08/2017
---


# TableBackground Object (PowerPoint)

Represents the background associated with a  **Table** object.


## Remarks

Use the  **[Background](table-background-property-powerpoint.md)** property of a **[Table](table-object-powerpoint.md)** object to return the **TableBackground** object associated with the table.

 To get a **Table** object from an existing shape, use the **Table** property of the **[Shape](shape-object-powerpoint.md)** or **[ShapeRange](shaperange-object-powerpoint.md)** object that contains the table. You can create a shape that contains a table by using the **[AddTable](shapes-addtable-method-powerpoint.md)** method of the **[Shapes](shapes-object-powerpoint.md)** collection.

The properties of the  **TableBackground** object return objects that represent various aspects of the formatting associated with a table.


- Use the  **[Fill](tablebackground-fill-property-powerpoint.md)** property to return a **[FillFormat](fillformat-object-powerpoint.md)** object.
    
- Use the  **[Picture](tablebackground-picture-property-powerpoint.md)** property to return a **[PictureFormat](pictureformat-object-powerpoint.md)** object.
    
- Use the  **[Reflection](tablebackground-reflection-property-powerpoint.md)** property to return an **[ReflectionFormat](http://msdn.microsoft.com/library/9684dbb3-5b99-113b-9808-1173fdd719a9%28Office.15%29.aspx)** object.
    
- Use the  **[Shadow](tablebackground-shadow-property-powerpoint.md)** property to return a **[ShadowFormat](shadowformat-object-powerpoint.md)** object.
    

## Example

The following example shows how to get a  **TableBackground** object and set two of its properties.


```vb
Public Sub TableBackground_Example() 
 
    Dim pptShape As PowerPoint.Shape 
    Dim pptTable As PowerPoint.Table 
    Dim pptTableBackground As PowerPoint.TableBackground 
    Dim pptFillFormat As PowerPoint.FillFormat 
     
    Set pptShape = ActivePresentation.Slides(2).Shapes.AddTable(3, 3) 
    Set pptTable = pptShape.Table 
    Set pptTableBackground = pptTable.Background 
    Set pptFillFormat = pptTableBackground.Fill 
     
    ' Add a patterned fill to the table background 
    pptFillFormat.Patterned (msoPatternSmallGrid) 
     
    ' Add a shadow to the table background 
    pptTableBackground.Shadow.Visible = msoTrue 
     
End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

