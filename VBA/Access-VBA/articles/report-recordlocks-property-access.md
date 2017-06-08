---
title: Report.RecordLocks Property (Access)
keywords: vbaac10.chm13696
f1_keywords:
- vbaac10.chm13696
ms.prod: access
api_name:
- Access.Report.RecordLocks
ms.assetid: 21f8d145-e417-a7a1-e697-b1e07434c760
ms.date: 06/08/2017
---


# Report.RecordLocks Property (Access)

You can use the  **RecordLocks** property to determine how records are locked and what happens when two users try to edit the same record at the same time. Read/write.


## Syntax

 _expression_. **RecordLocks**

 _expression_ A variable that represents a **Report** object.


## Remarks

When you edit a record, Microsoft Access can automatically lock that record to prevent other users from changing it before you are finished. For reports, the  **RecordLocks** property specifies whether records in the underlying table or query are locked while a report is previewed or printed.

The  **RecordLocks** property only applies to forms, reports, or queries in a Microsoft Access database.

The  **RecordLocks** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|No Locks|0|(Default) In reports, records aren't locked while the report is previewed or printed. In queries, records aren't locked while the query is run.|
|All Records|1|All records in the underlying table or query are locked while the report is previewed or printed, or while the query is run. Although users can read the records, the report has finished printing, or the query has finished running.|
|Edited Record|2|(Forms and queries only) |

 **Note**  Changing the  **RecordLocks** property of an open form or report causes an automatic recreation of the recordset.

You can use the No Locks setting for forms if only one person uses the underlying tables or queries or makes all the changes to the data.

In a multiuser database, you can use the No Locks setting if you want to use optimistic locking and warn users attempting to edit the same record on a form. You can use the Edited Record setting if you want to prevent two or more users editing data at the same time.

You can use the All Records setting when you need to ensure that no changes are made to data after you start to preview or print a report or run an append, delete, make-table, or update query.


## See also


#### Concepts


[Report Object](report-object-access.md)

