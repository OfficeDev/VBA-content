---
title: Return Objects from Collections
keywords: vbapp10.chm5193520
f1_keywords:
- vbapp10.chm5193520
ms.prod: powerpoint
ms.assetid: d81e1323-aa12-fa1a-aa75-3cc21d06c75f
ms.date: 06/08/2017
---


# Return Objects from Collections

The  **Item** method returns a single object from a collection. The following example sets the _firstPres_ variable to a [Presentation](presentation-object-powerpoint.md) object that represents presentation one variable to a [Presentation](presentation-object-powerpoint.md) object that represents presentation one.


```vb
Set firstPres = Presentations.Item(1)
```


The  **Item** method is the default method for most collections, so you can write the same statement more concisely by omitting the **Item** keyword.




```vb
Set firstPres = Presentations(1)
```

For more information about a specific collection, see the Help topic for that collection or the  **Item** method for the collection.

## Named Objects

Although you can usually specify an integer value by using the  **Item** method, it may be more convenient to return an object by name. Many objects are given automatically generated names when they are created. For example, the first slide you create will be automatically named "Slide1." If the first two shapes you create are a rectangle and an oval, their default names will be "Rectangle 1" and "Oval 2". You may want to give an object a more meaningful name to make it easier to refer to later. Most often, this is done by setting the object's **Name** property. The following example sets a meaningful name for a slide as it is added. You can then use the name instead of the index number to refer to the slide.


```vb
ActivePresentation.Slides.Add(1, 1).Name = "Home Page Slide"
With ActivePresentation.Slides("Home Page Slide")
    .FollowMasterBackground = False
    .Background.Fill.PresetGradient _
        msoGradientDiagonalDown, 1, msoGradientBrass
End With
```

 **Predefined Index Values**

Some collections have predefined index values you can use to return single objects. Each predefined index value is represented by a constant. For example, you specify a  **PpTextStyleType** constant with the **Item** method of the [TextStyles](textstyles-object-powerpoint.md) collection to return a single text style.

The following example sets the margins for the body area on slides in the active presentation.




```vb
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).TextFrame
    .MarginBottom = 50
    .MarginLeft = 50
    .MarginRight = 50
    .MarginTop = 50
End With
```


