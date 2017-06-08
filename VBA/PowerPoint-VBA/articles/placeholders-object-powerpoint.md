---
title: Placeholders Object (PowerPoint)
keywords: vbapp10.chm544000
f1_keywords:
- vbapp10.chm544000
ms.prod: powerpoint
api_name:
- PowerPoint.Placeholders
ms.assetid: d16e06e4-185a-1b99-52a7-4787a4990684
ms.date: 06/08/2017
---


# Placeholders Object (PowerPoint)

A collection of all the  **Shape** objects that represent placeholders on the specified slide.


## Remarks

 Each **Shape** object in the **Placeholders** collection represents a placeholder for text, a chart, a table, an organizational chart, or some other type of object. If the slide has a title, the title is the first placeholder in the collection.

You can delete individual placeholders by using the [Delete](shapenodes-delete-method-powerpoint.md)method, and you can restore deleted placeholders by using the [AddPlaceholder](shapes-addplaceholder-method-powerpoint.md)method, but you cannot add any more placeholders to a slide than it had when it was created. To change the number of placeholders on a given slide, set the [Layout](slide-layout-property-powerpoint.md)property.


## Example

Use the [Placeholders](shapes-placeholders-property-powerpoint.md)property to return the  **Placeholders** collection. Use **Placeholders** (index), where index is the placeholder index number, to return a **Shape** object that represents a single placeholder. Note that for any slide that has a title, `Shapes`.Title is equivalent to  `Shapes.Placeholders(1)`.The following example adds a new slide with a Bulleted List slide layout to the beginning of the presentation, sets the text for the title, and then adds two paragraphs to the text placeholder.


```vb
Set sObj = ActivePresentation.Slides.Add(1, ppLayoutText).Shapes
sObj.Title.TextFrame.TextRange.Text = "This is the title text"
sObj.Placeholders(2).TextFrame.TextRange.Text = _
    "Item 1" &; Chr(13) &; "Item 2"
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

