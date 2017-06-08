---
title: Export a Report to XML
ms.prod: access
ms.assetid: 7e746a40-6227-1481-f631-702c3cf42d0f
ms.date: 06/08/2017
---


# Export a Report to XML

This procedure exports the Invoice report in the current database to an XML file. It also exports presentation information, and places images in the Images folder. The procedure exports the report to the default HTML wrapper. In addition, it creates a file containing the ReportML list.


```vb
Private Sub ExportReport() 
 
 Const CREATE_REPORTML = 16 
 
 Application.ExportXML _ 
 ObjectType:=acExportReport, _ 
 DataSource:="Invoice", _ 
 DataTarget:="C:\Invoice.xml", _ 
 PresentationTarget:="C:\InvoiceReport.xsl", _ 
 ImageTarget:="C:\Images", _ 
 OtherFlags:=CREATE_REPORTML 
 
End Sub
```


