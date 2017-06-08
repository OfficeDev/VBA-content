---
title: Pane.ViewType Property (PowerPoint)
keywords: vbapp10.chm631005
f1_keywords:
- vbapp10.chm631005
ms.prod: powerpoint
api_name:
- PowerPoint.Pane.ViewType
ms.assetid: 6114b581-a9f5-a4b7-827e-99004fea4e58
ms.date: 06/08/2017
---


# Pane.ViewType Property (PowerPoint)

Returns the type of view for the specified pane. Read-only.


## Syntax

 _expression_. **ViewType**

 _expression_ A variable that represents a **Pane** object.


### Return Value

PpViewType


## Remarks

The value of the  **ViewType** property can be one of these **PpViewType** constants.


||
|:-----|
|**ppViewHandoutMaster**|
|**ppViewMasterThumbnails**|
|**ppViewNormal**|
|**ppViewNotesMaster**|
|**ppViewNotesPage**|
|**ppViewOutline**|
|**ppViewPrintPreview**|
|**ppViewSlide**|
|**ppViewSlideMaster**|
|**ppViewSlideSorter**|
|**ppViewThumbnails**|
|**ppViewTitleMaster**|

## Example

If the view in the active pane is slide view, this example makes the notes pane the active pane. The notes pane is the third member of the  **Panes** collection.


```vb
With ActiveWindow

    If .ActivePane.ViewType = ppViewSlide Then

        .Panes(3).Activate

    End If

End With
```


## See also


#### Concepts


[Pane Object](pane-object-powerpoint.md)

