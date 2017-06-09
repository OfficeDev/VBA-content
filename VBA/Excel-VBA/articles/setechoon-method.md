---
title: SetEchoOn Method
keywords: vbagr10.chm3077084
f1_keywords:
- vbagr10.chm3077084
ms.prod: excel
api_name:
- Excel.SetEchoOn
ms.assetid: 48490f33-63ef-aef1-8e54-51ac5d8f35e5
ms.date: 06/08/2017
---


# SetEchoOn Method

Returns a Chart object.

 _expression_. **SetEchoOn**( **_EchoOn_**)

 _expression_ Required. An expression that returns a **Chart** object.

 **EchoOn**Optional  **Variant**.

## Example

This example sets the echo on for the first object in the application.


```vb
Sub UseEchoOn() 
 
 Dim grpOne As Graph.Chart 
 
 Set grpOne = Application.ActiveSheet.OLEObjects(1).Object 
 
 grpOne.SetEchoOn 
 
End Sub
```


