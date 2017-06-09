---
title: Selection.LayoutChangeDirection Method (Visio)
keywords: vis_sdr.chm11162195
f1_keywords:
- vis_sdr.chm11162195
ms.prod: visio
api_name:
- Visio.Selection.LayoutChangeDirection
ms.assetid: 1c40348c-1884-1501-3609-aebf2e87686c
ms.date: 06/08/2017
---


# Selection.LayoutChangeDirection Method (Visio)

Revises the layout of a selection of connected shapes by rotating or flipping the connected shapes as a unit, without rotating or flipping the individual shapes.


## Syntax

 _expression_ . **LayoutChangeDirection**( **_Direction_** )

 _expression_ A variable that represents a **[Selection](selection-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[VisLayoutDirection](vislayoutdirection-enumeration-visio.md)**|The layout action to take. See Remarks for possible values.|

### Return Value

 **Nothing**


## Remarks

The  _Direction_ parameter must be one of the following **VisLayoutDirection** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLayoutDirRotateRight**|0|Rotates the selection 90 degrees clockwise.|
| **visLayoutDirRotateLeft**|1|Rotates the selection 90 degrees counterclockwise.|
| **visLayoutDirFlipVert**|2|Flips the selection vertically.|
| **visLayoutDirFlipHorz**|3|Flips the selection horizontally.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LayoutChangeDirection** method to flip a selection of connected shapes vertically, without flipping the individual shapes.


```vb
Public Sub SelectionLayoutChangeDirection_Example()
  Dim vsoSelection As Visio.Selection 
  Set vsoSelection = ActiveWindow.Selection 
  vsoSelection.LayoutChangeDirection (visLayoutDirFlipVert) 
End Sub
```


