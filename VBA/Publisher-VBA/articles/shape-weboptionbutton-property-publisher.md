---
title: Shape.WebOptionButton Property (Publisher)
keywords: vbapb10.chm2228343
f1_keywords:
- vbapb10.chm2228343
ms.prod: publisher
api_name:
- Publisher.Shape.WebOptionButton
ms.assetid: 0c43387c-0cb6-5d6f-68cb-d1883ce17243
ms.date: 06/08/2017
---


# Shape.WebOptionButton Property (Publisher)

Returns the  **[WebOptionButton](weboptionbutton-object-publisher.md)** object associated with the specified shape.


## Syntax

 _expression_. **WebOptionButton**

 _expression_A variable that represents a  **Shape** object.


### Return Value

WebOptionButton


## Example

This example creates a new Web option button and specifies that its default state is selected.


```vb
Dim shpNew As Shape 
Dim wobTemp As WebOptionButton 
 
Set shpNew = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlOptionButton, Left:=100, _ 
 Top:=123, Width:=16, Height:=10) 
 
Set wobTemp = shpNew.WebOptionButton 
 
wobTemp.Selected = msoTrue
```


