---
title: Selection.Distribute Method (Visio)
keywords: vis_sdr.chm11151420
f1_keywords:
- vis_sdr.chm11151420
ms.prod: visio
api_name:
- Visio.Selection.Distribute
ms.assetid: 7750167b-b4ef-c1b6-68f4-1f40ab1fd33e
ms.date: 06/08/2017
---


# Selection.Distribute Method (Visio)

Distributes three or more selected shapes at regular intervals on the drawing page. The order of selection is irrelevant.


## Syntax

 _expression_ . **Distribute**( **_Distribute_** , **_GlueToGuide_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Distribute_|Required| **VisDistributeTypes**|Specifies how the shapes are distributed. See Remarks for possible values.|
| _GlueToGuide_|Optional| **Boolean**|If  **True** , creates guides and glues selected shapes to them. If **False** , does not. Default is **False** .|

### Return Value

Nothing


## Remarks

The following possible values for  _Distribute_ are declared in **VisDistributeTypes** in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDistHorzCenter**|2|Distributes shapes horizontally so that their bottom edges are uniformly spaced.|
| **visDistHorzLeft**|1|Distributes shapes horizontally so that their left edges are uniformly spaced.|
| **visDistHorzRight**|3|Distributes shapes horizontally so that their right edges are uniformly spaced.|
| **visDistHorzSpace**|0|Distributes shapes horizontally so that there is a uniform space between shapes.|
| **visDistVertBottom**|7|Distributes shapes vertically so that their bottom edges are uniformly spaced.|
| **visDistVertMiddle**|6|Distributes shapes vertically so that their centers are uniformly spaced.|
| **visDistVertSpace**|4|Distributes shapes vertically so that there is a uniform space between shapes.|
| **visDistVertTop**|5|Distributes shapes vertically so that their top edges are uniformly spaced.|
Calling the  **Distribute** method is equivalent to setting options in the **Distribute Shapes** dialog box (on the **Home** tab, click **Position**, point to  **Space Shapes**, and then click  **More Distribute Options**). 

Passing  **True** for the optional _GlueToGuide_ argument is the equivalent of selecting the **Create guides and glue shapes to them** check box in the **Distribute Shapes** dialog box.

When you pass  **True** for _GlueToGuide_, Visio creates guides to retain the distribution of the shapes. You can select and move the outermost guides to move the shapes without changing their distribution.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Distribute** method to distribute three shapes vertically so that their right edges are uniformally spaced and glued to guides.


```vb
Public Sub Distribute_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoShape3 As Visio.Shape 
 
 Set vsoShape1 = Application.ActiveWindow.Page.DrawRectangle(1, 9, 3, 7) 
 Set vsoShape2 = Application.ActiveWindow.Page.DrawRectangle(3, 6, 5, 5) 
 Set vsoShape3 = Application.ActiveWindow.Page.DrawRectangle(6, 4, 8, 2) 
 
 ActiveWindow.DeselectAll 
 
 ActiveWindow.Select vsoShape1, visSelect 
 ActiveWindow.Select vsoShape2, visSelect 
 ActiveWindow.Select vsoShape3, visSelect 
 
 Application.ActiveWindow.Selection.Distribute visDistVertRight, True 
 
End Sub
```


