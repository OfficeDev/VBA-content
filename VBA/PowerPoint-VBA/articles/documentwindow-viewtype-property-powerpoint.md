---
title: DocumentWindow.ViewType Property (PowerPoint)
keywords: vbapp10.chm511006
f1_keywords:
- vbapp10.chm511006
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.ViewType
ms.assetid: 95eb4962-6d7a-41bd-fdae-757287f06350
ms.date: 06/08/2017
---


# DocumentWindow.ViewType Property (PowerPoint)

Returns or sets the type of the view contained in the specified document window. Read/write.


## Syntax

 _expression_. **ViewType**

 _expression_ A variable that represents a **DocumentWindow** object.


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

This example changes the view in the active window to slide sorter view if the window is currently displayed in normal view.


```vb
With Application.ActiveWindow

    If .ViewType = ppViewNormal Then

        .ViewType = ppViewSlideSorter

    End If

End With
```


## See also


#### Concepts



[DocumentWindow Object](documentwindow-object-powerpoint.md)

