---
title: Report.PageFooter Property (Access)
keywords: vbaac10.chm13698
f1_keywords:
- vbaac10.chm13698
ms.prod: access
api_name:
- Access.Report.PageFooter
ms.assetid: 82cd1c0f-2823-9b61-a1fd-66c02c6aaadf
ms.date: 06/08/2017
---


# Report.PageFooter Property (Access)

You can use the  **PageFooter** property to specify whether a report's page footer is printed on the same page as a report footer. Read/write **Byte**.


## Syntax

 _expression_. **PageFooter**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **PageFooter** property use the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|All Pages|0|(Default) The page footer is printed on all pages of a report.|
|Not With Rpt Hdr|1|The page footer isn't printed on the same page as the report header.|
|Not With Rpt Ftr|2|The page footer isn't printed on the same page as the report footer. Microsoft Access prints the report footer on a new page.|
|Not With Rpt Hdr/Ftr|3|The page footer isn't printed on a page that has either a report header or a report footer. Microsoft Access prints the report footer on a new page.|
You can set the  **PageFooter** property only in report Design view.

Microsoft Access normally prints report page footers on every page in a report, including the first and last.


 **Note**  When forms are printed, page footers are always printed on all pages.


## Example

The following example sets the  **PageFooter** property for a report to Not With Rpt Hdr. To run this example, type the following line in the Debug window for a report named Report1.


```vb
Reports!Report1.PageFooter = 1
```


## See also


#### Concepts


[Report Object](report-object-access.md)

