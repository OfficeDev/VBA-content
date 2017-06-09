---
title: Printer.DriverName Property (Access)
keywords: vbaac10.chm12860
f1_keywords:
- vbaac10.chm12860
ms.prod: access
api_name:
- Access.Printer.DriverName
ms.assetid: 7434f44a-8b55-1f21-e595-363327199037
ms.date: 06/08/2017
---


# Printer.DriverName Property (Access)

Returns a  **String** indicating the name of the driver used by the specified printer. Read-only.


## Syntax

 _expression_. **DriverName**

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

