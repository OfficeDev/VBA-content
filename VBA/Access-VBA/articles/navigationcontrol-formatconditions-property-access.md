---
title: NavigationControl.FormatConditions Property (Access)
keywords: vbaac10.chm11038
f1_keywords:
- vbaac10.chm11038
ms.prod: access
api_name:
- Access.NavigationControl.FormatConditions
ms.assetid: 20e921d6-e800-fc75-c93a-981815d694ab
ms.date: 06/08/2017
---


# NavigationControl.FormatConditions Property (Access)

You can use the  **FormatConditions** property to return a read-only reference to the **[FormatConditions](formatconditions-object-access.md)** collection and its related properties.


## Syntax

 _expression_. **FormatConditions**

 _expression_ A variable that represents a **NavigationControl** object.


## Example

The following example sets format properties for an existing conditional format for the "Textbox1" control.


```vb
With Forms("forms1").Controls("Textbox1").FormatConditions(1) 
 .BackColor = RGB(255,255,255) 
 .FontBold = True 
 .ForeColor = RGB(255,0,0) 
End With
```


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

