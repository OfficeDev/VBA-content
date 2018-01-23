---
title: Form.RecordsetClone Property (Access)
keywords: vbaac10.chm13496
f1_keywords:
- vbaac10.chm13496
ms.prod: access
api_name:
- Access.Form.RecordsetClone
ms.assetid: d73ef798-477d-9c36-6e29-82b22352c60b
ms.date: 06/08/2017
---


# Form.RecordsetClone Property (Access)

You can use the  **RecordsetClone** property to refer to a form's **Recordset** object specified by the form's **[RecordSource](form-recordsource-property-access.md)** property. Read-only.


## Syntax

 _expression_. **RecordsetClone**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **RecordsetClone** property setting is a copy of the underlying query or table specified by the form's **RecordSource** property. If a form is based on a query, for example, referring to the **RecordsetClone** property is the equivalent of cloning a **Recordset** object by using the same query. If you then apply a filter to the form, the **Recordset** object reflects the filtering.

This property is available only by using [Visual Basic](set-properties-by-using-visual-basic.md) and is read-only in all views.

You use the  **RecordsetClone** property to navigate or operate on a form's records independent of the form itself. For example, you can use the **RecordsetClone** property when you want to use a method, such as the DAO **Find** methods, that can't be used with forms.

When a new  **Recordset** object is opened, its first record is the current record. If you one of the **Find** method or one of the **Move** methods to make any other record in the **Recordset** object current, you must synchronize the current record in the **Recordset** object with the form's current record by assigning the value of the DAO **Bookmark** property to the form's **[Bookmark](form-bookmark-property-access.md)** property.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Transfer data from Access to Excel](http://www.utteraccess.com/forum/Transfer-data-Access-Ex-t1672619.html)
    

## Example

The following example uses the  **RecordsetClone** property to create a new clone of the **Recordset** object from the Orders form and then prints the names of the fields in the Immediate window.


```vb
Sub Print_Field_Names() 
    Dim rst As Recordset, intI As Integer 
    Dim fld As Field 
 
    Set rst = Me.RecordsetClone 
    For Each fld in rst.Fields 
        ' Print field names. 
        Debug.Print fld.Name 
    Next 
End Sub
```

The next example uses the  **RecordsetClone** property and the **Recordset** object to synchronize a recordset's record with the form's current record. When a company name is selected from a combo box, the **FindFirst** method is used to locate the record for that company and the **Recordset** object's DAO **Bookmark** property is assigned to the form's **Bookmark** property, causing the form to display the found record.




```vb
Sub SupplierID_AfterUpdate() 
    Dim rst As Recordset 
    Dim strSearchName As String 
 
    Set rst = Me.RecordsetClone 
    strSearchName = Str(Me!SupplierID) 
    rst.FindFirst "SupplierID = " &; strSearchName 
        If rst.NoMatch Then 
            MsgBox "Record not found" 
        Else 
            Me.Bookmark = rst.Bookmark 
        End If 
    rst.Close 
End Sub
```

You can use the  **RecordCount** property to count the number of records in a **Recordset** object. The following example shows how you can combine the **RecordCount** property and the **RecordsetClone** property to count the records in a form:




```vb
Forms!Orders.RecordsetClone.MoveLast 
MsgBox "My form contains " _ 
    &; Forms!Orders.RecordsetClone.RecordCount _ 
    &; " records.", vbInformation, "Record Count"
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Form Object](form-object-access.md)

