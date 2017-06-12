---
title: Application.Printers Property (Access)
keywords: vbaac10.chm12596
f1_keywords:
- vbaac10.chm12596
ms.prod: access
api_name:
- Access.Application.Printers
ms.assetid: 71383404-8244-6e9b-9c72-8963e0901901
ms.date: 06/08/2017
---


# Application.Printers Property (Access)

Returns the  **[Printers](printers-object-access.md)** collection representing all the available printers on the current system. Read-only **Printers** collection.


## Syntax

 _expression_. **Printers**

 _expression_ A variable that represents an **Application** object.


## Example

The following example displays information about all the printers available on the current system.


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


[Application Object](application-object-access.md)

