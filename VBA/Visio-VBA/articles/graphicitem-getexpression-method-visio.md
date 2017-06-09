---
title: GraphicItem.GetExpression Method (Visio)
keywords: vis_sdr.chm16960425
f1_keywords:
- vis_sdr.chm16960425
ms.prod: visio
api_name:
- Visio.GraphicItem.GetExpression
ms.assetid: 61864d97-a61b-549a-6f41-d741c19a330f
ms.date: 06/08/2017
---


# GraphicItem.GetExpression Method (Visio)

Gets the label of the shape data item (custom property) that the  **GraphicItem** represents, or the value of the expression string that is part of a **GraphicItem** object?s rule, against which shape data is evaluated.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetExpression**( **_Field_** , **_Expression_** )

 _expression_ An expression that returns a **GraphicItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required| **VisGraphicField**|Out parameter. A constant value from the  **VisGraphicField** enumeration specifying whether Expression is a shape-data-item label or the formula in the ShapeSheet spreadsheet of the primary data field (column) assigned to the **GraphicItem** . See Remarks for possible values.|
| _Expression_|Required| **String**|Out parameter. The ShapeSheet expression associated with the Field parameter.|

### Return Value

Nothing


## Remarks

The Field value returned as an out parameter is one of the following values from the  **VisGraphicField** enumeration, which is declared in the Microsoft Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGraphicPropertyLabel**|1|The label of a shape data item.|
| **visGraphicExpression**|2|The ShapeSheet formula of a shape data item.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetExpression** method to get the value of the expression string for a data graphic item. It gets the most recently applied expression for the first graphic item in the **GraphicItems** collection of the **Master** object of type **visTypeDataGraphic** (commonly called a data graphic) named "Data Graphic," and prints the field and the expression in the **Immediate** window.


 **Note**  You can determine the name of an existing data graphic master by moving your mouse over the master in the  **Data Graphics** task pane in the Visio user interface.

The macro assumes that a data graphic named "Data Graphic" exists in the current document. For more information about adding a data graphic master to the  **Masters** collection of the current document, see **[Masters.AddEx ](masters-addex-method-visio.md)** .




```vb
Public Sub GetExpression() 
 
    Dim vsoGraphicItem As Visio.GraphicItem 
    Set vsoGraphicItem = ActiveDocument.Masters("Data Graphic").GraphicItems(1) 
    Dim strExpression As String 
    Dim fieldName As VisGraphicField 
     
    vsoGraphicItem.GetExpression fieldName, strExpression 
    Debug.Print strExpression 
    Debug.Print fieldName 
 
End Sub
```


