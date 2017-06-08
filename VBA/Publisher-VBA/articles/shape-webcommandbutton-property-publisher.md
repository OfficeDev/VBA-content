---
title: Shape.WebCommandButton Property (Publisher)
keywords: vbapb10.chm2228340
f1_keywords:
- vbapb10.chm2228340
ms.prod: publisher
api_name:
- Publisher.Shape.WebCommandButton
ms.assetid: c20b937b-6f53-fdc1-830a-4044831c351a
ms.date: 06/08/2017
---


# Shape.WebCommandButton Property (Publisher)

Returns the  **[WebCommandButton](webcommandbutton-object-publisher.md)** object associated with the specified shape.


## Syntax

 _expression_. **WebCommandButton**

 _expression_A variable that represents a  **Shape** object.


### Return Value

WebCommandButton


## Example

This example creates a Web form Submit command button and sets the script path and file name to run when a user clicks the button.


```vb
Dim shpNew As Shape 
Dim wcbTemp As WebCommandButton 
 
Set shpNew = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36) 
 
Set wcbTemp = shpNew.WebCommandButton 
 
With wcbTemp 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" _ 
 &; "scripts/ispscript.cgi" 
End With
```


