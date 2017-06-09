---
title: AcExportXMLOtherFlags Enumeration (Access)
keywords: vbaac10.chm13251
f1_keywords:
- vbaac10.chm13251
ms.prod: access
api_name:
- Access.AcExportXMLOtherFlags
ms.assetid: ebc80f42-56e8-e024-241a-a2ddc5d752ca
ms.date: 06/08/2017
---


# AcExportXMLOtherFlags Enumeration (Access)

Use with the  **ExportXML** method to specify other behaviors associated with exporting to XML.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**acEmbedSchema**|1|Writes schema information into the document specified by the DataTarget argument; this value takes precedence over the SchemaTarget argument.|
|**acExcludePrimaryKeyAndIndexes**|2|Does not export primary key and index schema properties.|
|**acExportAllTableAndFieldProperties**|32|The exported schema contains properties of the table and its fields.|
|**acLiveReportSource**|8|Creates a live link to a remote Microsoft SQL Server 2000 database. Valid only when you are exporting reports that are bound to a Microsoft SQL Server 2000 database.|
|**acPersistReportML**|16|Persists the exported object's ReportML information.|
|**acRunFromServer**|4|Creates an Active Server Pages (ASP) wrapper; otherwise, default is an HTML wrapper. Applies only when you are exporting reports.|

