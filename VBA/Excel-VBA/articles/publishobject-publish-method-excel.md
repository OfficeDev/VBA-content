---
title: PublishObject.Publish Method (Excel)
keywords: vbaxl10.chm652074
f1_keywords:
- vbaxl10.chm652074
ms.prod: excel
api_name:
- Excel.PublishObject.Publish
ms.assetid: 3bb70102-c440-8e49-1734-d72945324d5c
ms.date: 06/08/2017
---


# PublishObject.Publish Method (Excel)

Saves an item or a collection of items in a document to a Web page.


## Syntax

 _expression_ . **Publish**( **_Create_** )

 _expression_ A variable that represents a **PublishObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Create_|Optional| **Variant**|If the HTML file exists, setting this argument to  **True** replaces the file, and setting this argument to **False** inserts the item or items at the end of the file. If the file does not exist, then the file is created regardless of the value of the _Create_ argument.|

## Remarks

The  **[FileName](publishobject-filename-property-excel.md)** property returns or sets the location and name of the HTML file.


## Example

This example saves the range D5:D9 on the First Quarter worksheet in the active workbook to a Web page named stockreport.htm. The Spreadsheet component is used to make the Web page interactive.


```vb
With ActiveWorkbook.PublishObjects.Add(xlSourceRange, _ 
 "\\Server1\sharedfolder\stockreport.htm", "First Quarter", _ 
 "$D$5:$D$9", xlHtmlStatic, "Book2_25082", "") 
 .Publish (True) 
 .AutoRepublish = True 
End With
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-excel.md)

