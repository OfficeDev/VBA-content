---
title: Printers Object (Access)
keywords: vbaac10.chm12881
f1_keywords:
- vbaac10.chm12881
ms.prod: access
api_name:
- Access.Printers
ms.assetid: 5200c507-75ae-f9a8-c737-c28e175e7ea4
ms.date: 06/08/2017
---


# Printers Object (Access)

The  **Printers** collection contains **[Printer](printer-object-access.md)** objects representing all the printers available on the current system.


## Remarks

Use the  **[Printers](application-printers-property-access.md)** property of the **Application** object to return the **Printers** collection. You can enumerate through the **Printers** collection by using the **For Each...Next** statement.

You can refer to an individual  **Printer** object in the **Printers** collection either by referring to the printer by name, or by referring to its index within the collection.

The  **Printers** collection is indexed beginning with zero. If you refer to a printer by its index, the first printer is Printers(0), the second printer is Printers(1), and so on.

You can't add or delete a  **Printer** object from the **Printers** collection.


## Example

The following example displays information about all the printers available to the system.


```
Dim prtLoop As Printer 
 
For Each prtLoop In Application.Printers 
 With prtLoop 
 MsgBox "Device name: " &amp; .DeviceName &amp; vbCr _ 
 &amp; "Driver name: " &amp; .DriverName &amp; vbCr _ 
 &amp; "Port: " &amp; .Port 
 End With 
Next prtLoop
```


## Properties



|**Name**|
|:-----|
|[Application](printers-application-property-access.md)|
|[Count](printers-count-property-access.md)|
|[Item](printers-item-property-access.md)|
|[Parent](printers-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
