---
title: TabStops Object (PowerPoint)
keywords: vbapp10.chm573000
f1_keywords:
- vbapp10.chm573000
ms.prod: powerpoint
api_name:
- PowerPoint.TabStops
ms.assetid: e23b36de-6a4d-84e5-bec1-8c3e0fd80c13
ms.date: 06/08/2017
---


# TabStops Object (PowerPoint)

A collection of all the  **[TabStop](tabstop-object-powerpoint.md)** objects on one ruler.


## Example

Use the [TabStops](ruler-tabstops-property-powerpoint.md)property to return the  **TabStops** collection. The following example clears all the tab stops for the text in shape two on slide one in the active presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2) _
        .TextFrame.Ruler.TabStops
    For t = .Count To 1 Step -1
        .Item(t).Clear
    Next
End With
```

Use the [Add](tabstops-add-method-powerpoint.md)method to create a tab stop and add it to the  **TabStops** collection. The following example adds a tab stop to the body-text style on the slide master for the active presentation. The new tab stop will be positioned 2 inches (144 points) from the left edge of the ruler and will be left aligned.




```vb
ActivePresentation.SlideMaster _
    .TextStyles(ppBodyStyle).Ruler.TabStops.Add ppTabStopLeft, 144
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

