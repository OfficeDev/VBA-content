---
title: Report.PageHeader Property (Access)
keywords: vbaac10.chm13697,vbaac10.chm5495
f1_keywords:
- vbaac10.chm13697,vbaac10.chm5495
ms.prod: access
api_name:
- Access.Report.PageHeader
ms.assetid: 9f9fe114-b5a5-39c7-d2c0-39453948ace6
ms.date: 06/08/2017
---


# Report.PageHeader Property (Access)

You can use the  **PageHeader** property to specify whether a report's page header is printed on the same page as a report header. Read/write **Byte**.


## Syntax

 _expression_. **PageHeader**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **PageHeader** property use the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|All Pages|0|(Default) The page header is printed on all pages of a report.|
|Not With Rpt Hdr|1|The page header isn't printed on the same page as the report header.|
|Not With Rpt Ftr|2|The page header isn't printed on the same page as the report footer. Microsoft Access prints the report footer on a new page.|
|Not With Rpt Hdr/Ftr|3|The page header isn't printed on a page that has either a report header or a report footer. Microsoft Access prints the report footer on a new page.|
You can set the  **PageHeader** property only in report Design view.

Microsoft Access normally prints report page headers on every page in a report, including the first and last.


 **Note**  When forms are printed, page headers are always printed on all pages.


## Example

The following example sets the  **PageHeader** property for a report to Not With Rpt Hdr. To run this example, type the following line in the Debug window for a report named Report1.


```vb
Reports!Report1.PageHeader = 1
```


## See also


#### Concepts


[Report Object](report-object-access.md)

