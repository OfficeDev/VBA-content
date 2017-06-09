---
title: Form.RecordSource Property (Access)
keywords: vbaac10.chm13345
f1_keywords:
- vbaac10.chm13345
ms.prod: access
api_name:
- Access.Form.RecordSource
ms.assetid: a473695a-7645-744d-bf69-760e1f2b9fb1
ms.date: 06/08/2017
---


# Form.RecordSource Property (Access)

You can use the  **RecordSource** property to specify the source of the data for a form. Read/write **String**.


## Syntax

 _expression_. **RecordSource**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **RecordSource** property setting can be a table name, a query name, or an SQL statement. For example, you can use the following settings.



|**Sample setting**|**Description**|
|:-----|:-----|
|Employees|A table name specifying the Employees table as the source of data.|
| `SELECT Orders!OrderDate FROM Orders;`|An SQL statement specifying the OrderDate field on the Orders table as the source of data. You can bind a control on the form or report to the OrderDate field in the Orders table by setting the control's  **ControlSource** property to OrderDate.|

 **Note**  Changing the record source of an open form or report causes an automatic requery of the underlying data. If a form's  **Recordset** property is set at runtime, the form's **RecordSource** property is updated.

After you have created a form or report, you can change its source of data by changing the  **RecordSource** property. The **RecordSource** property is also useful if you want to create a reusable form or report. For example, you could create a form that incorporates a standard design, then copy the form and change the **RecordSource** property to display data from a different table, query, or SQL statement.

Limiting the number of records contained in a form's record source can enhance performance, especially when your application is running on a network. For example, you can set a form's  **RecordSource** property to an SQL statement that returns a single record and change the form's record source depending on criteria selected by the user.


## Example

The following example sets a form's  **RecordSource** property to the Customers table:


```vb
Forms!frmCustomers.RecordSource = "Customers"
```

The next example changes a form's record source to a single record in the Customers table, depending on the company name selected in the  `cmboCompanyName` combo box control. The combo box is filled by an SQL statement that returns the customer ID (in the bound column) and the company name. The CustomerID has a Text data type.




```vb
Sub cmboCompanyName_AfterUpdate() 
 Dim strNewRecord As String 
 strNewRecord = "SELECT * FROM Customers " _ 
 &; " WHERE CustomerID = '" _ 
 &; Me!cmboCompanyName.Value &; "'" 
 Me.RecordSource = strNewRecord 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

