---
title: GraphicItem.SetExpression Method (Visio)
keywords: vis_sdr.chm16960430
f1_keywords:
- vis_sdr.chm16960430
ms.prod: visio
api_name:
- Visio.GraphicItem.SetExpression
ms.assetid: e0fd9a38-1fc0-3189-9def-64f2c181951d
ms.date: 06/08/2017
---


# GraphicItem.SetExpression Method (Visio)

Sets the value of the expression string that is part of a  **GraphicItem** object?s rule, against which shape data (custom properties) are evaluated.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **SetExpression**( **_Field_** , **_Expression_** )

 _expression_ An expression that returns a **GraphicItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required| **VisGraphicField**|The label or formula of the primary data field (column) assigned to the  **GraphicItem** . See Remarks for possible values.|
| _Expression_|Required| **String**|The ShapeSheet expression associated with the Field parameter.|

### Return Value

Nothing


## Remarks

The Field parameter should be one of the following values from the  **VisGraphicField** enumeration, which is declared in the Microsoft Visio Type Library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGraphicPropertyLabel**|1|The label of a shape data item.|
| **visGraphicExpression**|2|The ShapeSheet formula of a shape data item.|
When you pass the  **SetExpression** method a shape-data-item label (that is, when you pass **visGraphicPropertyLabel** for the Field parameter), you must enclose the label within curly braces ({}). For example, if you want to pass the name of the "Cost" shape-data item, you must write it like this: {Cost}.

You can reference the shape data of a shape other than the one to which the data graphic is applied by passing the name of the shape followed by an exclamation point (!) and then the name of the field. For example, in the example shown below, to refer to the width of the shape named Ellipse.34, you could use the following syntax:




```
vsoGraphicItem.SetExpression visGraphicExpression, "Ellipse.34!Width"
```

Before you can edit a graphic item, including setting its expression string, you must use the  **[Master.Open](master-open-method-visio.md)** method to open for editing a copy of the data graphic master whose **GraphicItems** collection the graphic item belongs to. After you have set the expression of the graphic item and made whatever other edits you want to make, use the **Master.Close** method to commit changes.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SetExpression** method to set the value of the expression string for a data graphic item. It opens a copy of the **Master** object of type **visTypeDataGraphic** (commonly called a data graphic) named "Data Graphic," and then sets the expression of the first graphic item in the **GraphicItems** collection of the data graphic to display the width of any shape to which the data graphic is applied.


 **Note**  You can determine the name of an existing data graphic master by moving your mouse over the master in the  **Data Graphics** task pane in the Visio user interface.

Then it closes the master and uses the  **GetExpression** method to get the mostly recently applied expression for the graphic item. Finally, it prints the field type and the value of the expression in the **Immediate** window.

The macro assumes that a data graphic named "Data Graphic" exists in the current document. For more information about adding a data graphic master to the  **Masters** collection of the current document, see **[Masters.AddEx ](masters-addex-method-visio.md)** .




```vb
Public Sub SetExpression_Example() 
 
    Dim vsoMaster As Visio.Master 
    Dim vsoMasterCopy As Visio.Master 
    Dim vsoGraphicItem As Visio.GraphicItem 
    Dim strExpression As String 
    Dim fieldType As VisGraphicField 
 
    Set vsoMaster = Visio.ActiveDocument.Masters("Data Graphic") 
    Set vsoMasterCopy = vsoMaster.Open 
    Set vsoGraphicItem = vsoMasterCopy.GraphicItems(1) 
       
    vsoGraphicItem.SetExpression visGraphicExpression, "Width" 
    vsoMasterCopy.Close 
     
    vsoMaster.GraphicItems(1).GetExpression fieldType, strExpression 
     
    Debug.Print "Field type is "; fieldType 
    Debug.Print "Expression is "; strExpression 
     
End Sub
```


