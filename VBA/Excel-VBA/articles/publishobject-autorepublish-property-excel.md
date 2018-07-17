---
title: PublishObject.AutoRepublish Property (Excel)
keywords: vbaxl10.chm652082
f1_keywords:
- vbaxl10.chm652082
ms.prod: excel
api_name:
- Excel.PublishObject.AutoRepublish
ms.assetid: edf5579f-eb70-85d3-de5d-1ae229359898
ms.date: 06/08/2017
---


# PublishObject.AutoRepublish Property (Excel)

When a workbook is saved, Microsoft Excel determines if any item in the  **[PublishObjects](publishobjects-object-excel.md)** collection has the **AutoRepublish** property set to **True** and, if so, republishes it. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **AutoRepublish**

 _expression_ A variable that represents a **PublishObject** object.


## Example

This example publishes a range on a worksheet to an HTML file on the C: drive. When the user saves the workbook containing the worksheet, Excel will automatically republish the range to the same HTML file. This example assumes that the user has read/write access to the web page and that cells A1 through D10 in the worksheet have values in them.


```vb
Sub PublishToWeb() 
 
 With ActiveWorkbook.PublishObjects.Add( _ 
 SourceType:= xlSourceRange, _ 
 Filename:="C:\Work.htm", _ 
 Sheet:="Sheet1", _ 
 Source:="A1:D10", _ 
 HtmlType:=xlHtmlStatic, _ 
 DivID:="Book1.xls_130489") 
 .Publish 
 .AutoRepublish = True 
 End With 
 
End Sub
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-excel.md)

