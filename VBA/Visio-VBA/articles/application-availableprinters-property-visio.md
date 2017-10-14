---
title: Application.AvailablePrinters Property (Visio)
keywords: vis_sdr.chm10050510
f1_keywords:
- vis_sdr.chm10050510
ms.prod: visio
api_name:
- Visio.Application.AvailablePrinters
ms.assetid: bd070ee3-4f32-1ff0-427c-d61b7778e6c5
ms.date: 06/08/2017
---


# Application.AvailablePrinters Property (Visio)

Returns a list of installed printers. Read-only.


## Syntax

 _expression_ . **AvailablePrinters**

 _expression_ A variable that represents an **Application** object.


### Return Value

String()


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AvailablePrinters** property to get a list of available printers.


```vb
Public Sub AvailablePrinters_example() 
 
 Dim aStrPrinters() As String 
 Dim strPrinter As Variant 
 
 aStrPrinters = Application.AvailablePrinters 
 
 For Each strPrinter In aStrPrinters 
 
 Debug.Print strPrinter 
 
 Next 
 
End Sub
```


