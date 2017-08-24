---
title: InstalledPrinters Object (Publisher)
keywords: vbapb10.chm8978431
f1_keywords:
- vbapb10.chm8978431
ms.prod: publisher
api_name:
- Publisher.InstalledPrinters
ms.assetid: 8cf9b194-70bc-7963-6a08-d08401d4b6f3
ms.date: 06/08/2017
---


# InstalledPrinters Object (Publisher)

Represents the collection of all  **Printer** objects, each of which represents one of the printers installed on the computer.
 


## Remarks

To provide the user a choice of printers to print a publication, you can iterate through the  **InstalledPrinters** collection to get a list of the names of all the printers installed on the computer, as shown in the following example.
 

 
The default property of the  **InstalledPrinters** collection is **Item**.
 

 

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **[PrinterName](printer-printername-property-publisher.md)** and **[IsActivePrinter](printer-isactiveprinter-property-publisher.md)** properties of the **Printer** object to get a list of all the installed printers on the computer and to determine which of them is currently the active printer.
 

 

```
Public Sub InstalledPrinters_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 Debug.Print pubPrinter.PrinterName 
 If pubPrinter.IsActivePrinter Then 
 Debug.Print "This is the active printer." 
 End If 
 Next 
 
End Sub 

```


## Properties



|**Name**|
|:-----|
|[Application](installedprinters-application-property-publisher.md)|
|[Count](installedprinters-count-property-publisher.md)|
|[Item](installedprinters-item-property-publisher.md)|
|[Parent](installedprinters-parent-property-publisher.md)|

