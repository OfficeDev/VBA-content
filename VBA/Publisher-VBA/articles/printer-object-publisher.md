---
title: Printer Object (Publisher)
keywords: vbapb10.chm9043967
f1_keywords:
- vbapb10.chm9043967
ms.prod: publisher
api_name:
- Publisher.Printer
ms.assetid: 46f8c6a2-4cf1-bb6a-1214-a751440870f2
ms.date: 06/08/2017
---


# Printer Object (Publisher)

A  **Printer** object represents a printer installed on your computer.


## Remarks

Many of the properties, such as  **PaperSize**, **PaperSource**, and **PaperOrientation**, of the **Printer** object correspond to the settings in the **Print Setup** dialog box ( **File** menu) in the Microsoft Publisher user interface .

The collection of all the printers installed on your computer is represented by the  **InstalledPrinters** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **PrinterName** and **IsActivePrinter** properties of the **Printer** object to get a list of all the installed printers on the computer, determine which of them is currently the active printer, and get some of the settings of the active printer. The macro displays the results in the **Immediate** window.


```
Public Sub Printer_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 Debug.Print pubPrinter.PrinterName 
 If pubPrinter.IsActivePrinter Then 
 Debug.Print "This is the active printer" 
 Debug.Print "Paper size is ", pubPrinter.PaperSize 
 Debug.Print "Paper orientation is ", pubPrinter.PaperOrientation 
 Debug.Print "Paper source is ", pubPrinter.PaperSource 
 End If 
 Next 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c7eadef4-8206-7e86-b0fe-3c3fe7d07f25%28Office.15%29.aspx)|
|[DriverType](http://msdn.microsoft.com/library/99c3b4e5-a55a-0f8d-3767-d035d9d6e4df%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/2030a3d4-2e42-679c-6084-7a3959271e58%28Office.15%29.aspx)|
|[IsActivePrinter](http://msdn.microsoft.com/library/578fc5d4-2601-66db-cdec-657814756e29%28Office.15%29.aspx)|
|[IsColor](http://msdn.microsoft.com/library/ae466c89-8da0-986b-c3f8-b0aea651dffe%28Office.15%29.aspx)|
|[IsDuplex](http://msdn.microsoft.com/library/d39beb76-8a30-5f2d-3f04-016cfac943fa%28Office.15%29.aspx)|
|[PaperHeight](http://msdn.microsoft.com/library/2c97adb8-0a24-c375-6105-375b203d5640%28Office.15%29.aspx)|
|[PaperOrientation](http://msdn.microsoft.com/library/f57986b6-e6c4-7a47-af93-56036d667240%28Office.15%29.aspx)|
|[PaperSize](http://msdn.microsoft.com/library/fa7962fb-3ca0-470a-2337-3193ed0be2aa%28Office.15%29.aspx)|
|[PaperSource](http://msdn.microsoft.com/library/3c3f9007-c1ea-6957-6fa5-b34873e0a17f%28Office.15%29.aspx)|
|[PaperWidth](http://msdn.microsoft.com/library/e2f0392f-56b2-0ccb-c96c-0bccf2bfe0a0%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4f8994d4-423e-8cc6-fb8f-50c47659e892%28Office.15%29.aspx)|
|[PrintableRect](http://msdn.microsoft.com/library/9d5b8264-9213-3d89-0613-421a4872c158%28Office.15%29.aspx)|
|[PrinterName](http://msdn.microsoft.com/library/6987b89b-a77e-03c5-bd7e-015510034550%28Office.15%29.aspx)|
|[PrintMode](http://msdn.microsoft.com/library/47ca11d1-d058-0f4e-dd22-ec452dafaf1a%28Office.15%29.aspx)|

