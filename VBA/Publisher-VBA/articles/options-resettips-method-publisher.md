---
title: Options.ResetTips Method (Publisher)
keywords: vbapb10.chm1048616
f1_keywords:
- vbapb10.chm1048616
ms.prod: publisher
api_name:
- Publisher.Options.ResetTips
ms.assetid: a119aacc-ba19-f430-e8af-6d84c438ec25
ms.date: 06/08/2017
---


# Options.ResetTips Method (Publisher)

Resets tippages so that a user can view them when using features that have been used before.


## Syntax

 _expression_. **ResetTips**

 _expression_A variable that represents an  **Options** object.


## Remarks

The  **ResetTips** method is equivalent to clicking **Reset Tips** on the **User Assistance** tab of the **Options** dialog box ( **Tools** menu).


## Example

This example resets tip balloons.


```vb
Sub ResetTippages() 
 Options.ResetTips 
End Sub
```


