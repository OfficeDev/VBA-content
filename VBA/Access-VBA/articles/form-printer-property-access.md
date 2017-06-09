---
title: Form.Printer Property (Access)
keywords: vbaac10.chm13523
f1_keywords:
- vbaac10.chm13523
ms.prod: access
api_name:
- Access.Form.Printer
ms.assetid: c533271a-c500-57de-f16c-ed384698f829
ms.date: 06/08/2017
---


# Form.Printer Property (Access)

Returns or sets a  **[Printer](printer-object-access.md)** object representing the default printer on the current system. Read/write.


## Syntax

 _expression_. **Printer**

 _expression_ A variable that represents a **Form** object.


## Example

The following example makes the first printer in the  **[Printers](printers-object-access.md)** collection the default printer for the system, and then reports its name, driver information, and port information.


```vb
Dim prtDefault As Printer 
 
Set Application.Printer = Application.Printers(0) 
 
Set prtDefault = Application.Printer 
 
With prtDefault 
 MsgBox "Device name: " &; .DeviceName &; vbCr _ 
 &; "Driver name: " &; .DriverName &; vbCr _ 
 &; "Port: " &; .Port 
End With 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

