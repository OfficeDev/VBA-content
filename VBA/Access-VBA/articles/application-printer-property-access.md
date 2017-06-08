---
title: Application.Printer Property (Access)
keywords: vbaac10.chm12597
f1_keywords:
- vbaac10.chm12597
ms.prod: access
api_name:
- Access.Application.Printer
ms.assetid: a8398360-f11c-72b9-4b71-7b042889ac9c
ms.date: 06/08/2017
---


# Application.Printer Property (Access)

Returns or sets a  **[Printer](printer-object-access.md)** object representing the default printer on the current system. Read/write.


## Syntax

 _expression_. **Printer**

 _expression_ A variable that represents an **Application** object.


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


[Application Object](application-object-access.md)

