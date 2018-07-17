---
title: Printer.DeviceName Property (Access)
keywords: vbaac10.chm12859
f1_keywords:
- vbaac10.chm12859
ms.prod: access
api_name:
- Access.Printer.DeviceName
ms.assetid: bf4acead-26b9-603d-2ead-537822913405
ms.date: 06/08/2017
---


# Printer.DeviceName Property (Access)

Returns a  **String** indicating name of the specified printer device. Read-only.


## Syntax

 _expression_. **DeviceName**

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

