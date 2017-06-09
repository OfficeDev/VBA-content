---
title: Saving Documents as Web Pages
keywords: vbaxl10.chm5285317
f1_keywords:
- vbaxl10.chm5285317
ms.prod: excel
ms.assetid: ea07da4e-39f6-d04e-00cc-d52eb87f652f
ms.date: 06/08/2017
---


# Saving Documents as Web Pages

In Microsoft Excel, you can save a workbook, worksheet, chart, range, query table, PivotTable report, print area, or AutoFilter range to a Web page. You can also edit HTML files directly in Excel.


## Saving a Document as a Web Page

Saving a document as a Web page is the process of creating and saving an HTML file and any supporting files. To do this, use the  **[SaveAs](workbook-saveas-method-excel.md)** method, as shown in the following example, which saves the active workbook as C:\Reports\myfile.htm.


```vb
ActiveWorkbook.SaveAs _ 
 Filename:="C:\Reports\myfile.htm", _ 
 FileFormat:=xlHTML
```


## Customizing the Web Page

You can customize the appearance, content, browser support, editing support, graphics formats, screen resolution, file organization, and encoding of the HTML document by setting properties of the  **[DefaultWebOptions](defaultweboptions-object-excel.md)** object and the  **[WebOptions](weboptions-object-excel.md)** object. The  **DefaultWebOptions** object contains application-level properties. These settings are overridden by any workbook-level property settings that have the same names (these are contained in the **WebOptions** object).

After setting the attributes, you can use the  **[Publish](publishobject-publish-method-excel.md)** method to save the workbook, worksheet, chart, range, query table, PivotTable report, print area, or AutoFilter range to a Web page. The following example sets various application-level properties and then sets the  **[AllowPNG](weboptions-allowpng-property-excel.md)** property of the active workbook, overriding the application-level default setting. Finally, the example saves the range as "C:\Reports\1998_Q1.htm."




```vb
With Application.DefaultWebOptions 
 .RelyonVML = True 
 .AllowPNG = True 
 .PixelsPerInch = 96 
End With 
With ActiveWorkbook 
 .WebOptions.AllowPNG = False 
 With .PublishObjects(1) 
 .FileName = "C:\Reports\1998_Q1.htm" 
 .Publish 
 End With 
End With
```

You can also save the files directly to a Web server. The following example saves a range to a Web server, giving the Web page the URL address http://example.homepage.com/annualreport.htm.




```vb
With ActiveWorkbook 
 With .WebOptions 
 .RelyonVML = True 
 .PixelsPerInch = 96 
 End With 
 With .PublishObjects(1) 
 .FileName = _ 
 "http://example.homepage.com/annualreport.htm" 
 .Publish 
 End With 
End With
```


## Opening an HTML Document in Microsoft Excel

To edit an HTML document in Excel, first open the document by using the  **[Open](workbooks-open-method-excel.md)** method. The following example opens the file "C:\Reports\1997_Q4.htm" for editing.


```vb
Workbooks.Open Filename:="C:\Reports\1997_Q4.htm"
```

After opening the file, you can customize the appearance, content, browser support, editing support, graphics formats, screen resolution, file organization, and encoding of the HTML document by setting properties of the  **DefaultWebOptions** and **WebOptions** objects.


