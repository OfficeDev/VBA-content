---
title: PublishObject.DivID Property (Excel)
keywords: vbaxl10.chm652075
f1_keywords:
- vbaxl10.chm652075
ms.prod: excel
api_name:
- Excel.PublishObject.DivID
ms.assetid: a60bb113-e04f-7de7-98f0-3ddb0e51dcdc
ms.date: 06/08/2017
---


# PublishObject.DivID Property (Excel)

Returns the unique identifier used for identifying an HTML <DIV> tag on a Web page. The tag is associated with an item in a document that you have saved to a Web page. An item can be an entire workbook, a worksheet, a selected print range, an AutoFilter range, a range of cells, a chart, a PivotTable report, or a query table. Read-only  **String** .


## Syntax

 _expression_ . **DivID**

 _expression_ A variable that represents a **PublishObject** object.


## Example

This example saves a range of cells to a Web page, and then it obtains the identifier from the <DIV> tag of this item and finds the line on the saved Web page (q198.htm). The example also creates a copy of the Web page (newq1.htm) and inserts a comment line before the <DIV> tag in the copy of the file.


```vb
Set objPO = ActiveWorkbook.PublishObjects.Add( _ 
 SourceType:=xlSourceRange, _ 
 Filename:="\\Server1\Reports\q198.htm", _ 
 Sheet:="Sheet1", _ 
 Source:="C2:D6", _ 
 HtmlType:=xlHtmlStatic) 
objPO.Publish 
strTargetDivID = objPO.DivID 
Open "\\Server1\Reports\q198.htm" For Input As #1 
Open "\\Server1\Reports\newq1.htm" For Output As #2 
While Not EOF(1) 
 Line Input #1, strFileLine 
 If InStr(strFileLine, strTargetDivID) > 0 And _ 
 InStr(strFileLine, "<div") > 0 Then 
 Print #2, "<!--Saved item-->" 
 End If 
 Print #2, strFileLine 
Wend 
Close #2 
Close #1
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-excel.md)

