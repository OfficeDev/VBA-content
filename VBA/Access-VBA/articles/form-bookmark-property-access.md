---
title: Form.Bookmark Property (Access)
keywords: vbaac10.chm13421
f1_keywords:
- vbaac10.chm13421
ms.prod: access
api_name:
- Access.Form.Bookmark
ms.assetid: e214a924-9110-a3de-9812-b9ec5cbad8ed
ms.date: 06/08/2017
---


# Form.Bookmark Property (Access)

You can use the  **Bookmark** property with forms to set a bookmark that uniquely identifies a particular record in the form's underlying table, query, or SQL statement. Read/write **Variant**.


## Syntax

 _expression_. **Bookmark**

 _expression_ A variable that represents a **Form** object.


## Remarks




 **Note**  You get or set the form's  **Bookmark** property separately from the ADO **Bookmark** or DAO **Bookmark** property of the underlying table or query.

When a bound form is opened in Form view, each record is assigned a unique bookmark. In Visual Basic, you can save the bookmark for the current record by assigning the value of the form's  **Bookmark** property to a string variable. To return to a saved record after moving to a different record, set the form's **Bookmark** property to the value of the saved string variable. You can use the **StrComp** function to compare a **Variant** or string variable to a bookmark, or when comparing a bookmark against a bookmark. The third argument for the **StrComp** function must be set to a value of zero.




 **Note**  Bookmarks are not saved with the records they represent and are only valid while the form is open. They are re-created by Microsoft Access each time a bound form is opened.

There is no limit to the number of bookmarks you can save if each is saved with a unique string variable.

The  **Bookmark** property is only available for the form's current record. To save a bookmark for a record other than the current record, move to the desired record and assign the value of the **Bookmark** property to a string variable that identifies this record.

You can use bookmarks in any form that is based entirely on Microsoft Access tables. However, other database products may not support bookmarks. For example, you can't use bookmarks in a form based on a linked table that has no primary index.

Requerying a form invalidates any bookmarks set on records in the form. However, clicking  **Refresh** on the **Records** menu doesn't affect bookmarks.

Since Microsoft Access creates a unique bookmark for each record in a form's recordset when a form is opened, a form's bookmark will not work on another recordset, even when the two recordsets are based on the same table, query, or SQL statement. For example, suppose you open a form bound to the Customers table. If you then open the Customers table by using Visual Basic and use the ADO  **Seek** or DAO **Seek** method to locate a specific record in the table, you can't set the form's **Bookmark** property to the current table record. To perform this kind of operation you can use the ADO **Find** method or DAO **Find** methods with the form's **[RecordsetClone](form-recordsetclone-property-access.md)** property.

An error occurs if you set the  **Bookmark** property to a string variable and then try to return to that record after the record has been deleted.

The value of the  **Bookmark** property isn't the same as a record number.


## Example

To test the following example with the Northwind sample database, you need to add a command button named  `cmdFindContactName` to the Suppliers form, and then add the following code to the button's Click event. When the button is clicked, the user is asked to enter a portion of the contact name to find. If the name is found, the form's **Bookmark** property is set to the **Recordset** object's DAO **Bookmark** property, which moves the form's current record to the found name.


```vb
Private Sub cmdFindContactName_Click() 
 
 Dim rst As DAO.Recordset 
 Dim strCriteria As String 
 
 strCriteria = "[ContactName] Like '*" &; InputBox("Enter the " _ 
 &; "first few letters of the name to find") &; "*'" 
 
 Set rst = Me.RecordsetClone 
 rst.FindFirst strCriteria 
 If rst.NoMatch Then 
 MsgBox "No entry found.", vbInformation 
 Else 
 Me.Bookmark = rst.Bookmark 
 End If 
 
 Set rst = Nothing 
 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

