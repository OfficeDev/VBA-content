---
title: Report.DateGrouping Property (Access)
keywords: vbaac10.chm13699
f1_keywords:
- vbaac10.chm13699
ms.prod: access
api_name:
- Access.Report.DateGrouping
ms.assetid: e2495aa7-06e9-8eaf-81d8-182c7d51559c
ms.date: 06/08/2017
---


# Report.DateGrouping Property (Access)

You can use the  **DateGrouping** property to specify how you want to group dates in a report. Read/write **Byte**.


## Syntax

 _expression_. **DateGrouping**

 _expression_ A variable that represents a **Report** object.


## Remarks

For example, using the US Defaults setting will cause the week to begin on Sunday. If you set a Date/Time field's  **[GroupOn](grouplevel-groupon-property-access.md)** property to Week, the report will group dates from Sunday to Saturday.

The  **DateGrouping** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|US Defaults|0|Microsoft Access uses the U.S. settings for the first day of the week (Sunday) and the first week of the year (starts on January 1).|
|Use System Settings|1|(Default) Microsoft Access uses settings based on the locale selected in the  **Regional Options** dialog box in Windows Control Panel.|

 **Note**  The  **DateGrouping** property setting applies to the entire report, not to a particular group in the report.

You can set the  **DateGrouping** property only in report Design view or in the **[Open](report-open-event-access.md)** event procedure of a report.

The sort order used in a report isn't affected by the  **DateGrouping** property setting.


## See also


#### Concepts


[Report Object](report-object-access.md)

