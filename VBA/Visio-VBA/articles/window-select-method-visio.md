---
title: Window.Select Method (Visio)
keywords: vis_sdr.chm11616530
f1_keywords:
- vis_sdr.chm11616530
ms.prod: visio
api_name:
- Visio.Window.Select
ms.assetid: 04394905-0b6b-a07d-4085-a46cecf8afe3
ms.date: 06/08/2017
---


# Window.Select Method (Visio)

Selects or clears an object.


## Syntax

 _expression_ . **Select**( **_SheetObject_** , **_SelectAction_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SheetObject_|Required| **[IVSHAPE]**|An expression that returns a  **Shape** object to select or clear.|
| _SelectAction_|Required| **Integer**|The type of selection action to take.|

### Return Value

Nothing


## Remarks

When used with the  **Window** object, the **Select** method will affect the selection in the Microsoft Visio window. The **Selection** object, however, is independent of the selection in the window. Therefore, using the **Select** method with a **Selection** object only affects the state of the object in memory; the Visio window is unaffected.

The following constants declared by the Visio type library in  **VisSelectArgs** show valid values for selection types.



|** Constant**|** Value**|**Description**|
|:-----|:-----|:-----|
| **visDeselect**|1|Cancels the selection of a shape but leaves the rest of the selection unchanged.|
| **visSelect**|2|Selects a shape but leaves the rest of the selection unchanged.|
| **visSubSelect**|3|Selects a shape whose parent is already selected.|
| **visSelectAll**|4|Selects a shape and all its peers.|
| **visDeselectAll**|256|Cancels the selection of a shape and all its peers.|
If SelectAction is  **visSubSelect** , the parent shape of SheetObject must already be selected.

You can combine  **visDeselectAll** with **visSelect** and **visSubSelect** to clear all shapes prior to selecting or subselecting other shapes.

If the object being operated on is a  **Selection** object, and if the **Select** method selects a **Shape** object whose **ContainingShape** property is different from the **ContainingShape** property of the **Selection** object, the **Select** method clears everything, even if the selection type value does not specify canceling the selection.

If the object being operated on is a  **Window** object, and if SelectAction is not **visSubSelect** , the parent shape of SheetObject must be the same shape as that returned by the **ContainingShape** property of the **Window.Selection** object.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVWindow.Select(Microsoft.Office.Interop.Visio.Shape, short)**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to select, clear, and subselect shapes.


```vb
 
Public Sub Select_Example()  
 
    Const MAX_SHAPES = 6  
    Dim vsoShapes(1 To MAX_SHAPES) As Visio.Shape  
    Dim intCounter As Integer  
 
    'Draw six rectangles.  
    For intCounter = 1 To MAX_SHAPES  
        Set vsoShapes(intCounter) = ActivePage.DrawRectangle(intCounter, intCounter + 1, intCounter + 1, intCounter)  
    Next intCounter  
 
    'Cancel the selection of all the shapes on the page.  
    ActiveWindow.DeselectAll  
 
     'Create a Selection object. 
    Dim vsoSelection As Visio.Selection  
    Set vsoSelection = ActiveWindow.Selection  
 
    'Select the first three shapes on the page. 
    For intCounter = 1 To 3  
        vsoSelection.Select vsoShapes(intCounter), visSelect  
    Next intCounter  
 
    'Group the selected shapes.  
    'Although the first three shapes are now grouped, the  
    'array vsoShapes() still contains them. 
    Dim vsoGroup As Visio.Shape  
    Set vsoGroup = vsoSelection.Group 
 
    'There are now four shapes on the page?a group that contains three  
    'subshapes, and three ungrouped shapes. Subselection is  
    'accomplished by selecting the parent shape first or one of the  
    'group's shapes already subselected.  
 
    'Select parent (group) shape. 
    ActiveWindow.Select vsoGroup, visDeselectAll + visSelect  
 
    'Subselect two of the shapes in the group. 
    ActiveWindow.Select vsoShapes(1), visSubSelect  
    ActiveWindow.Select vsoShapes(3), visSubSelect  
 
     'At this point two shapes are subselected, but we want to  
    'start a new selection that includes the last two shapes  
    'added to the page and the group. 
 
    'Note that the subselections that were made in the group  
    'are canceled by selecting another shape that is 
    'at the same level as the parent of the subselected shapes.  
 
    'Select just one shape. 
     ActiveWindow.Select vsoShapes(MAX_SHAPES), _  
        visDeselectAll + visSelect  
 
    'Select another shape. 
    ActiveWindow.Select vsoShapes(MAX_SHAPES - 1), visSelect  
 
    'Select the group.  
    ActiveWindow.Select vsoGroup, visSelect  
 
    'Select all but one shape on the page.  
    ActiveWindow.SelectAll  
    ActiveWindow.Select vsoShapes(MAX_SHAPES - 1), visDeselect  
 
End Sub
```


