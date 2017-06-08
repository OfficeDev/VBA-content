---
title: SlideRange Object (PowerPoint)
keywords: vbapp10.chm532000
f1_keywords:
- vbapp10.chm532000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange
ms.assetid: 440ab59d-744a-209f-bf28-d0acd3a21e1a
ms.date: 06/08/2017
---


# SlideRange Object (PowerPoint)

A collection that represents a notes page or a slide range, which is a set of slides that can contain as little as a single slide or as much as all the slides in a presentation. 


## Remarks

You can include whichever slides you want — chosen from all the slides in the presentation or from all the slides in the selection — to construct a slide range. For example, you could construct a  **SlideRange** collection that contains the first three slides in a presentation, all the selected slides in the presentation, or all the title slides in the presentation.

Just as you can work with several slides at the same time in the user interface by selecting them and applying a command, you can work with several slides at the same time programmatically by constructing a  **SlideRange** collection and applying properties or methods to it. And just as some commands in the user interface that work on single slides aren't valid when multiple slides are selected, some properties and methods that work on a **Slide** object or on a **SlideRange** collection that contains only one slide will fail if they're applied to a **SlideRange** collection that contains more than one slide. In general, if you can't do something manually when more than one slide is selected (such as return the individual shapes on one of the slides), you can't do it programmatically by using a **SlideRange** collection that contains more than one slide.

For those operations that work in the user interface whether you have a single slide or multiple slides selected (such as copying the selection to the Clipboard or setting the slide background fill), the associated properties and methods will work on a  **SlideRange** collection that contains more than one slide. Here are some general guidelines for how these properties and methods behave when they're applied to multiple slides.


- Applying a method to a  **SlideRange** collection is equivalent to applying the method to all the **Slide** objects in that range as a group.
    
- Setting the value of a property of the  **SlideRange** collection is equivalent to setting the value of the property in each slide in that range individually (for a property that takes an enumerated type, setting the value to the "Mixed" value has no effect).
    
- A property of the  **SlideRange** collection that returns an enumerated type returns the value of the property for an individual slide in the collection if all slides in the collection have the same value for that property. If the slides in the collection don't all have the same value for the property, the property returns the "Mixed" value.
    
- A property of the  **SlideRange** collection that returns a simple data type (such as **Long**, **Single**, or **String** ) returns the value of the property for an individual slide in the collection if all slides in the collection have the same value for that property. If the slides in the collection don't all have the same value for the property, the property will return - 2 or generate an error. For example, using the **Name** property on a **SlideRange** object that contains multiple slides will generate an error because each slide has a different value for its **Name** property.
    
- Some formatting properties of slides aren't set by properties and methods that apply directly to the  **SlideRange** collection, but by properties and methods that apply to an object contained in the **SlideRange** collection, such as the **ColorScheme** object. If the contained object represents operations that can be performed on multiple objects in the user interface, you'll be able to return the object from a **SlideRange** collection that contains more than one slide, and its properties and methods will follow the preceding rules. For example, you can use the **ColorScheme** property to return the **ColorScheme** object that represents the color schemes used on all the slides in the specified **SlideRange** collection. Setting properties for this **ColorScheme** object will also set these properties for the **ColorScheme** objects on all the individual slides in the **SlideRange** collection.
    
The following examples describe how to:


- Return a set of slides that you specify by name or index number
    
- Return all or some of the selected slides in a presentation
    
- Return a notes page
    
- Apply properties and methods to a slide range
    

## Example

Use  **Slides.Range** (index), where index is the name or index number of the slide or an array that contains either names or index numbers of slides, to return a **SlideRange** collection that represents a set of slides in a presentation. You can use the **Array** function to construct an array of names or index numbers. The following example sets the background fill for slides one and three in the active presentation.


```vb
With ActivePresentation.Slides.Range(Array(1, 3))

    .FollowMasterBackground = False
    .Background.Fill.PresetGradient msoGradientHorizontal, _
         1, msoGradientLateSunset

End With
```

The following example sets the background fill for the slides named "Intro" and "Big Chart" in the active presentation. Note that slides are assigned automatically generated names of the form Sliden (where n is an integer) when they're created. To assign a more meaningful name to a slide, use the [Name](slide-name-property-powerpoint.md)property.




```vb
With ActivePresentation.Slides.Range(Array("Intro", "Big Chart"))

    .FollowMasterBackground = False
    .Background.Fill.PresetGradient msoGradientHorizontal, _
        1, msoGradientLateSunset

End With
```

Although you can use the [Range](slides-range-method-powerpoint.md)method to return any number of slides, it is simpler to use the [Item](slides-item-method-powerpoint.md)method if you only want to return a single member of the  **SlideRange** collection. For example, `Slides(1)` is simpler than `Slides.Range(1)`.

Use the [SlideRange](selection-sliderange-property-powerpoint.md)property of the  **[Selection](selection-object-powerpoint.md)** object to return all the slides in the selection. The following example sets the background fill for all the selected slides in window one, assuming that there's at least one slide selected.




```vb
With Windows(1).Selection.SlideRange

    .FollowMasterBackground = False
    .Background.Fill.PresetGradient msoGradientHorizontal, _
        1, msoGradientLateSunset

End With
```

Use  **Selection.SlideRange** (index), where index is the slide name or index number, to return a single slide from the selection. The following example sets the background fill for slide two in the collection of selected slides in window one, assuming that there are at least two slides selected.




```vb
With Windows(1).Selection.SlideRange(2)

    .FollowMasterBackground = False
    .Background.Fill.PresetGradient msoGradientHorizontal, _
        1, msoGradientLateSunset

End With
```

Use the  **NotesPage** property to return a **SlideRange** collection that represents the specified notes page. The following example inserts text into placeholder two (the notes area) on the notes page for slide one in the active presentation.




```vb
ActivePresentation.Slides(1).NotesPage.Shapes _
    .Placeholders(2).TextFrame.TextRange.InsertAfter "Added Text"
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

