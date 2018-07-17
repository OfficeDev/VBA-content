---
title: Form.RecordLocks Property (Access)
keywords: vbaac10.chm13362
f1_keywords:
- vbaac10.chm13362
ms.prod: access
api_name:
- Access.Form.RecordLocks
ms.assetid: 9080f7dd-259e-8b13-9648-3269bc7321d3
ms.date: 06/08/2017
---


# Form.RecordLocks Property (Access)

You can use the  **RecordLocks** property to determine how records are locked and what happens when two users try to edit the same record at the same time. Read/write.


## Syntax

 _expression_. **RecordLocks**

 _expression_ A variable that represents a **Form** object.


## Remarks

When you edit a record, Microsoft Access can automatically lock that record to prevent other users from changing it before you are finished. For forms, the  **RecordLocks** property specifies how records in the underlying table or query are locked when data in a multiuser database is updated

The  **RecordLocks** property only applies to forms, reports, or queries in a Microsoft Access database.

The  **RecordLocks** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|No Locks|0|(Default) In forms, two or more users can edit the same record simultaneously. This is also called "optimistic" locking. If two users attempt to save changes to the same record, Microsoft Access displays a message to the user who tries to save the record second. This user can then discard the record, copy the record to the Clipboard, or replace the changes made by the other user. This setting is typically used on read-only forms or in single-user databases. It is also used in multiuser databases to permit more than one user to be able to make changes to the same record at the same time.|
|All Records|1|All records in the underlying table or query are locked while the form is open in Form view or Datasheet view. Although users can read the records, no one can edit, add, or delete any records until the form is closed.|
|Edited Record|2|(Forms and queries only) A page of records is locked as soon as any user starts editing any field in the record and stays locked until the user moves to another record. Consequently, a record can be edited by only one user at a time. This is also called "pessimistic" locking.|

 **Note**  Changing the  **RecordLocks** property of an open form or report causes an automatic recreation of the recordset.

You can use the No Locks setting for forms if only one person uses the underlying tables or queries or makes all the changes to the data.

In a multiuser database, you can use the No Locks setting if you want to use optimistic locking and warn users attempting to edit the same record on a form. You can use the Edited Record setting if you want to prevent two or more users editing data at the same time.

In Form view or Datasheet view, each locked record has a locked indicator in its record selector.

To change the default  **RecordLocks** property setting for forms, click **Options** on the **Tools** menu, click the **Advanced** tab on the **Options** dialog box, and then select the desired option under **Default record locking**.

Data in a form, report, or query from an Open Database Connectivity (ODBC) database is treated as if the No Locks setting were chosen, regardless of the  **RecordLocks** property setting.


## Example

The following example sets the  **RecordLocks** property of the "Employees" form to Edited Record (a page of records is locked as soon as any user starts editing any field in the record and stays locked until the user moves to another record).


```vb
Forms("Employees").RecordLocks = 2
```


## See also


#### Concepts


[Form Object](form-object-access.md)

