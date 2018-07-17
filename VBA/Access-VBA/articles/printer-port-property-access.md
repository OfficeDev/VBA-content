---
title: Printer.Port Property (Access)
keywords: vbaac10.chm12865
f1_keywords:
- vbaac10.chm12865
ms.prod: access
api_name:
- Access.Printer.Port
ms.assetid: 0fef85fb-fbe7-eada-1629-d56b6008e039
ms.date: 06/08/2017
---


# Printer.Port Property (Access)

Returns a  **String** indicating the port name of the specified printer. Read-only.


## Syntax

 _expression_. **Port**

 _expression_ A variable that represents a **Printer** object.


## Example

The following example displays information about all the printers available to the system.


```vb
Dim prtLoop As Printer 
 
For Each prtLoop In Application.Printers 
 With prtLoop 
 MsgBox "Device name: " &; .DeviceName &; vbCr _ 
 &; "Driver name: " &; .DriverName &; vbCr _ 
 &; "Port: " &; .Port 
 End With 
Next prtLoop
```


## See also


#### Concepts


[Printer Object](printer-object-access.md)

