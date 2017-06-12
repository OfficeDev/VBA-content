---
title: Selection Object (Publisher)
keywords: vbapb10.chm917503
f1_keywords:
- vbapb10.chm917503
ms.prod: publisher
api_name:
- Publisher.Selection
ms.assetid: 1ebee88b-a39e-ea3a-48b0-6205621853af
ms.date: 06/08/2017
---


# Selection Object (Publisher)

Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the publication, or it represents the cursor if nothing in the publication is selected. There can only be one  **Selection** object per publication window pane, and only one **Selection** object in the entire application can be active.
 


## Example

Use the  **[Selection](document-selection-property-publisher.md)** property to return the **Selection** object. If no object qualifier is used with the **Selection** property, Microsoft Publisher returns the selection from the active pane of the active publication window. The following example copies the current selection from the active publication.
 

 

```
Sub CopySelection() 
 Selection.ShapeRange.Copy 
End Sub
```

The following example determines what type of item is selected and, if it is an autoshape, fills the first shape in the selection with color. This example assumes there is at least one item selected in the active pubication.
 

 



```
Sub SelectedShape() 
 If Selection.Type = pbSelectionShape Then 
 Selection.ShapeRange.Item(1).Fill.ForeColor _ 
 .RGB = RGB(Red:=200, Green:=20, Blue:=255) 
 End If 
End Sub
```

The following example copies the selection and pastes it into the first shape on the second page of the active publication.
 

 



```
Sub CopyPasteSelection() 
 Selection.TextRange.Copy 
 With ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
 .Collapse Direction:=pbCollapseEnd 
 .InsertAfter NewText:=vbLf 
 .Paste 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Unselect](selection-unselect-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](selection-application-property-publisher.md)|
|[ChildShapeRange](selection-childshaperange-property-publisher.md)|
|[Parent](selection-parent-property-publisher.md)|
|[ShapeRange](selection-shaperange-property-publisher.md)|
|[TableCellRange](selection-tablecellrange-property-publisher.md)|
|[TextRange](selection-textrange-property-publisher.md)|
|[Type](selection-type-property-publisher.md)|

