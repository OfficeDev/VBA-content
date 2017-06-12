---
title: ImportExportSpecification Object (Access)
keywords: vbaac10.chm13327
f1_keywords:
- vbaac10.chm13327
ms.prod: access
api_name:
- Access.ImportExportSpecification
ms.assetid: a274faba-6da3-35c5-52fc-3341e8def24a
ms.date: 06/08/2017
---


# ImportExportSpecification Object (Access)

Represents a saved import or export operation.


## Remarks

A  **ImportExportSpecification** object contains all the information Access needs to repeat an import or export operation without your having to provide any input. For example, an import specification that imports data from a Microsoft Office Excel 2007 workbook stores the name of the source Excel file, the name of the destination database, and other details, such as whether you appended to or created a new table, primary key information, field names, and so on.

Use the  **[Add](http://msdn.microsoft.com/library/c048c45f-15e9-6347-b953-c9a5702d2bc5%28Office.15%29.aspx)** method of the **[ImportExportSpecifications](http://msdn.microsoft.com/library/9ddb9b30-36f3-5efb-8b15-69762c660338%28Office.15%29.aspx)** collection to create a new **ImportExportSpecification** object.

Use the  **[Execute](http://msdn.microsoft.com/library/fcb7cfd3-0c66-f441-9b58-1c6982125f98%28Office.15%29.aspx)** method to run saved import or export operation.


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/cc91c51e-1b2e-1d6e-b236-61a538843ce4%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/fcb7cfd3-0c66-f441-9b58-1c6982125f98%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/6ae597a1-8d1f-dc4f-c170-a4e664011a58%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/fa6f45a9-7358-3baa-12ad-e9ca46dd2104%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/365dffd4-295a-4db9-b31c-003890d94e0a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d2dc8f33-08fe-2b3b-178e-65c06cb25922%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/91799e23-304a-f2d9-9c22-779b79ab4700%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[ImportExportSpecification Object Members](http://msdn.microsoft.com/library/f170c0ad-07ab-f567-c75e-f35cca22f189%28Office.15%29.aspx)
