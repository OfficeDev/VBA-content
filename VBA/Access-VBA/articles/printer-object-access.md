---
title: Printer Object (Access)
keywords: vbaac10.chm12880
f1_keywords:
- vbaac10.chm12880
ms.prod: access
api_name:
- Access.Printer
ms.assetid: fba3eb15-db93-943a-421c-291761e7fa2b
ms.date: 06/08/2017
---


# Printer Object (Access)

A  **Printer** object corresponds to a printer available on your system.


## Remarks

A  **Printer** object is a member of the **[Printers](printers-object-access.md)** collection.

To return a reference to a particular  **Printer** object in the **Printers** collection, use any of the following syntax forms.



|**Syntax**|**Description**|
|:-----|:-----|
|**Printers** ! _devicename_|The  _devicename_ argument is the name of the **Printer** object as returned by the **DeviceName** property.|
|**Printers** (" _devicename_")|The  _devicename_ argument is the name of the **Printer** object as returned by the **DeviceName** property.|
|**Printers** ( _index_)|The  _index_ argument is the numeric position of the object within the collection. The valid range is from 0 to `Printers.Count-1`.|
You can use the properties of the  **Printer** object to set the printing characteristics for any of the printers available on your system.

Use the  **ColorMode**, **Copies**, **Duplex**, **Orientation**, **PaperBin**, **PaperSize**, and **PrintQuality** properties to specify print settings for a particular printer.

Use the  **LeftMargin**, **RightMargin**, **TopMargin**, **BottomMargin**, **ColumnSpacing**, **RowSpacing**, **DataOnly**, **DefaultSize**, **ItemLayout**, **ItemsAcross**, **ItemSizeHeight**, and **ItemSizeWidth** properties to specify how Microsoft Access should format the appearance of data on printed pages.

Use the  **DeviceName**, **DriverName**, and **Port** properties to return system information about a particular printer.


## Example

The following example displays system information about the first printer in the  **Printers** collection.


```
Dim prtFirst As Printer 
 
Set prtFirst = Application.Printers(0) 
 
With prtFirst 
 MsgBox "Device name: " &amp; .DeviceName &amp; vbCr _ 
 &amp; "Driver name: " &amp; .DriverName &amp; vbCr _ 
 &amp; "Port: " &amp; .Port 
End With
```


## Properties



|**Name**|
|:-----|
|[BottomMargin](printer-bottommargin-property-access.md)|
|[ColorMode](printer-colormode-property-access.md)|
|[ColumnSpacing](printer-columnspacing-property-access.md)|
|[Copies](printer-copies-property-access.md)|
|[DataOnly](printer-dataonly-property-access.md)|
|[DefaultSize](printer-defaultsize-property-access.md)|
|[DeviceName](printer-devicename-property-access.md)|
|[DriverName](printer-drivername-property-access.md)|
|[Duplex](printer-duplex-property-access.md)|
|[ItemLayout](printer-itemlayout-property-access.md)|
|[ItemsAcross](printer-itemsacross-property-access.md)|
|[ItemSizeHeight](printer-itemsizeheight-property-access.md)|
|[ItemSizeWidth](printer-itemsizewidth-property-access.md)|
|[LeftMargin](printer-leftmargin-property-access.md)|
|[Orientation](printer-orientation-property-access.md)|
|[PaperBin](printer-paperbin-property-access.md)|
|[PaperSize](printer-papersize-property-access.md)|
|[Port](printer-port-property-access.md)|
|[PrintQuality](printer-printquality-property-access.md)|
|[RightMargin](printer-rightmargin-property-access.md)|
|[RowSpacing](printer-rowspacing-property-access.md)|
|[TopMargin](printer-topmargin-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
