---
title: Page.ShapeAdded Event (Visio)
keywords: vis_sdr.chm10919225
f1_keywords:
- vis_sdr.chm10919225
ms.prod: visio
api_name:
- Visio.Page.ShapeAdded
ms.assetid: bc66eadc-21bc-7f17-6878-fddd9aaff855
ms.date: 06/08/2017
---


# Page.ShapeAdded Event (Visio)

Occurs after one or more shapes are added to a document.


## Syntax

Private Sub  _expression_ _**ShapeAdded**( **_ByVal Shape As [IVSHAPE]_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape or group of shapes that was added to the document.|

## Remarks

A  **Shape** object can serve as the source object for the **ShapeAdded** event if the shape's **Type** property is **visTypeGroup** (2) or **visTypePage** (1).

The  **SelectionAdded** and **ShapeAdded** events are similar in that they both fire after shape(s) are created. They differ in how they behave when a single operation adds several shapes. Suppose a **Paste** operation creates three new shapes. The **ShapeAdded** event fires three times and acts on each of the three objects. The **SelectionAdded** event fires once, and it acts on a **Selection** object in which the three new shapes are selected.

To determine if a  **ShapeAdded** event was triggered by a new shape or group of shapes being added to the page, by a set of existing shapes being grouped, or by a paste action, you can use the **Application.IsInScope** property. If **IsInScope** returns **True** when passed **visCmdObjectGroup** , the **ShapeAdded** event was triggered by a grouping action. If **IsInScope** returns **True** when passed **visCmdUFEditPaste** or **visCmdEditPasteSpecial** , the **ShapeAdded** event was triggered by a paste operation. If **IsInScope** returns **False** when passed all of these arguments, the event must have been triggered by new shapes being added to the page.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


## Example

This VBA example shows how to count shapes added to a drawing that are based on a master called  **Square**. Paste the code into the active document's project in Visio.

The  **DocumentSaved** event handler runs when the active document is saved. The handler initializes an integer variable, _intNumberOfSquares_ , which is used to store the count.

The  **ShapeAdded** event handler runs each time a shape is added to the drawing page, whether the shape is dragged from a stencil, drawn with a drawing tool, or pasted from the Clipboard. The handler checks the **Master** property of the new shape and, if the shape is based on the **Square** master, increments _intNumberOfSquares_ .




```vb
 
Dim intNumberOfSquares As Integer 
 
Private Sub Document_DocumentSaved(ByVal vsoDocument As Visio.IVDocument) 
 
 'Initialize number of squares added. 
 intNumberOfSquares = 0 
 
End Sub 
 
Private Sub Document_ShapeAdded(ByVal vsoShape As Visio.IVShape) 
 
 Dim vsoMaster As Visio.Master 
 
 'Get the Master property of the shape. 
 Set vsoMaster = vsoShape.Master 
 
 'Check whether the shape has a master. If not, 
 'the shape was created locally. 
 If Not (vsoMaster Is Nothing) Then 
 
 'Check whether the master is "Square". 
 If vsoMaster.Name = "Square" Then 
 
 'Increment the count for the number of squares added. 
 intNumberOfSquares = intNumberOfSquares + 1 
 
 End If 
 
 End If 
 
 MsgBox "Number of squares: " &; intNumberOfSquares, vbInformation, _ 
 "Document Created Example" 
 
End Sub
```


