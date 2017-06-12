---
title: Work with Shapes (Drawing Objects)
keywords: vbapp10.chm5194108
f1_keywords:
- vbapp10.chm5194108
ms.prod: powerpoint
ms.assetid: 3ffaaaea-6406-262b-2bc7-788699175266
ms.date: 06/08/2017
---


# Work with Shapes (Drawing Objects)

Shapes, or drawing objects, are represented by three different objects: the  **Shapes** collection, the **ShapeRange** collection, and the **Shape** object. In general, you use the **Shapes** collection to create shapes and when you want to iterate through all the shapes on a slide; you use the **Shape** object when you want to modify a single shape; and you use the **ShapeRange** collection when you want to modify multiple shapes the same way you can work with multiple selected shapes in the user interface.


## Setting properties for a shape

Many formatting properties of shapes are not set by properties that apply directly to the  **Shape** or **ShapeRange** object. Instead, related shape attributes are grouped under secondary objects, such as the **FillFormat** object, which contains all the properties that relate to the shape's fill, or the **LinkFormat** object, which contains all the properties that are unique to linked OLE objects. To set properties for a shape, you must first return the object that represents the set of related shape attributes and then set properties of that returned object. For example, you use the **Fill** property to return the **FillFormat** object, and then you set the **ForeColor** property of the **FillFormat** object to set the fill foreground color for the specified shape, as shown in the following example.


```vb
Set myDocument = ActivePresentation.Slides(1)
myDocument.Shapes(1).Fill.ForeColor.RGB = RGB(255, 0, 0)
```


## Applying a property or method to several shapes at the same time

In the user interface, there are some operations you can perform with several shapes selected; for example, you can select several shapes and set all their individual fills at once. There are other operations you can only perform with a single shape selected; for example, you can only edit the text in a shape if a single shape is selected.

In Visual Basic, there are two ways to apply properties and methods to a set of shapes. These two ways allow you to perform any operation that you can perform on a single shape on a range of shapes, whether or not you can perform the same operation in the user interface.


- If the operation works on multiple selected shapes in the user interface, you can perform the same operation in Visual Basic by constructing a  **ShapeRange** collection that contains the shapes you want to work with, and applying the appropriate properties and methods directly to the **ShapeRange** collection.
    
- If the operation does not work on multiple selected shapes in the user interface, you can still perform the operation in Visual Basic by looping through the  **Shapes** collection or through a **ShapeRange** collection that contains the shapes you want to work with, and applying the appropriate properties and methods to the individual **Shape** objects in the collection.
    
Many properties and methods that apply to the  **Shape** object and **ShapeRange** collection fail if applied to certain kinds of shapes. For example, the **TextFrame** property fails if applied to a shape that cannot contain text. If you are not positive that each shape in a **ShapeRange** collection can have a certain property or method applied to it, do not apply the property or method to the **ShapeRange** collection. If you want to apply one of these properties or methods to a collection of shapes, you must loop through the collection and test each individual shape to make sure it is an appropriate type of shape before applying the property or method to it.


## Applying a property or method to a ShapeRange collection

If you can perform an operation on multiple selected shapes in the user interface at the same time, you can do the programmatic equivalent by constructing a  **ShapeRange** collection and then applying the appropriate properties or methods to it. The following example constructs a shape range that contains the AutoShapes named "Big Star" and "Little Star" on `myDocument` and applies a gradient fill to them and applies a gradient fill to them.


```vb
Set myDocument = ActivePresentation.Slides(1) 
Set myRange = myDocument.Shapes _ 
    .Range(Array("Big Star", "Little Star")) 
myRange.Fill.PresetGradient msoGradientHorizontal, _ 
    1, msoGradientBrass
```

The following are general guidelines for how properties and methods behave when they are applied to a  **ShapeRange** collection.


- Applying a method to the collection is equivalent to applying the method to each individual  **Shape** object in that collection.
    
- Setting the value of a property of the collection is equivalent to setting the value of the property of each individual shape in that range.
    
- A property of the collection that returns a constant returns the value of the property for an individual shape in the collection if all shapes in the collection have the same value for that property. If not all shapes in the collection have the same value for the property, it returns the "mixed" constant.
    
- A property of the collection that returns a simple data type (such as  **Long**, **Single**, or **String** ) returns the value of the property for an individual shape if all shapes in the collection have the same value for that property.
    
- The value of some properties can be returned or set only if there is exactly one shape in the collection. If there is more than one shape in the collection, a run-time error occurs. This is generally the case for returning or setting properties when the equivalent action in the user interface is possible only with a single shape (actions such as editing text in a shape or editing the points of a freeform).
    
The preceding guidelines also apply when you are setting properties of shapes that are grouped under secondary objects of the  **ShapeRange** collection, such as the **FillFormat** object. If the secondary object represents operations that can be performed on multiple selected objects in the user interface, you will be able to return the object from a **ShapeRange** collection and set its properties. For example, you can use the **Fill** property to return the **FillFormat** object that represents the fills of all the shapes in the **ShapeRange** collection. Setting the properties of this **FillFormat** object will set the same properties for all the individual shapes in the **ShapeRange** collection.


## Looping through a Shapes or ShapeRange collection

Even if you cannot perform an operation on several shapes in the user interface at the same time by selecting them and then using a command, you can perform the equivalent action programmatically by looping through the  **Shapes** collection or through a **ShapeRange** collection that contains the shapes you want to work with, and applying the appropriate properties and methods to the individual **Shape** objects in the collection. The following example loops through all the shapes on `myDocument` and adds text to each shape that is an AutoShape. and adds text to each shape that is an AutoShape. `Set myDocument = ActivePresentation.Slides(1) For Each sh In myDocument.Shapes If sh.Type = msoAutoShape Then sh.TextFrame.TextRange.InsertAfter " (version 1)" End If Next`

The following example constructs  **ShapeRange** collection that contains all the currently selected shapes in the active window and sets the text in each shape in the collection that can contain text.




```vb
For Each sh in ActiveWindow.Selection.ShapeRange
    If sh.HasTextFrame Then
        sh.TextFrame.TextRange = "Initially selected"
    End If
Next

```


## Aligning, distributing, and grouping shapes in a shape range

Use the  **Align** and **Distribute** methods to position a set of shapes relative to one another or relative to the document that contains them. Use the **Group** method or the **Regroup** method to form a single grouped shape from a set of shapes.


