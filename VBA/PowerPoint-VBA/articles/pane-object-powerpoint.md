---
title: Pane Object (PowerPoint)
keywords: vbapp10.chm631000
f1_keywords:
- vbapp10.chm631000
ms.prod: powerpoint
api_name:
- PowerPoint.Pane
ms.assetid: 27862fd6-897d-893d-d5a8-b1e40b1b9d48
ms.date: 06/08/2017
---


# Pane Object (PowerPoint)

An object representing one of the three panes in normal view or the single pane of any other view in the document window.


## Remarks

Use  **Panes** (index), where index is the index number for a pane, to return a single **Pane** object. The following table lists the names of the panes in normal view with their corresponding index numbers.



|**Pane**|**Index number**|
|:-----|:-----|
|Outline|1|
|Slide|2|
|Notes|3|
When using a document window view other than normal view, use  **Panes** (1) to reference the single **Pane** object.

Use the [Activate](pane-activate-method-powerpoint.md)method to make the specified pane active.

Use the [ViewType](pane-viewtype-property-powerpoint.md)property to determine which pane is active. 

Normal view is the only view with multiple panes. All other document window views have only a single pane, which is the document window.


## Example

The following example uses the  **ViewType** property to determine whether the slide pane is the active pane. If it is, then the **Activate** method makes the notes pane the active pane.


```vb
With ActiveWindow

    If .ActivePane.ViewType = ppViewSlide Then

        .Panes(3).Activate

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

