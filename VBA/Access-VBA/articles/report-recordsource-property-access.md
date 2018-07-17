---
title: Report.RecordSource Property (Access)
keywords: vbaac10.chm13689
f1_keywords:
- vbaac10.chm13689
ms.prod: access
api_name:
- Access.Report.RecordSource
ms.assetid: aa3b31cc-21a6-5d56-8361-9fc232ffae97
ms.date: 11/30/2017
---


# Report.RecordSource Property (Access)

You can use the **RecordSource** property to specify the source of the data for a report. Read/write **String**.


## Syntax

_expression_. **RecordSource**

_expression_ A variable that represents a **Report** object.


## Remarks

The **RecordSource** property setting can be a table name, a query name, or an SQL statement. For example, you can use the following settings.

|**Sample setting**|**Description**|
|:-----|:-----|
|Employees|A table name specifying the Employees table as the source of data.|
| `SELECT Orders!OrderDate FROM Orders;`|An SQL statement specifying the  **OrderDate** field on the Orders table as the source of data. You can bind a control on the form or report to the **OrderDate** field in the Orders table by setting the control's **ControlSource** property to **OrderDate**.|

> [!NOTE]
> Changing the record source of an open form or report causes an automatic requery of the underlying data. If a form's **Recordset** property is set at runtime, the form's **RecordSource** property is updated.

After you have created a form or report, you can change its source of data by changing the **RecordSource** property. The **RecordSource** property is also useful if you want to create a reusable form or report. For example, you could create a form that incorporates a standard design, then copy the form and change the **RecordSource** property to display data from a different table, query, or SQL statement.


## Example

The following example sets a form's **RecordSource** property to the Customers table:

```vb
Forms!frmCustomers.RecordSource = "Customers"
```

<br/>

The next example changes a form's record source to a single record in the Customers table, depending on the company name selected in the `cmboCompanyName` combo box control. The combo box is filled by an SQL statement that returns the customer ID (in the bound column) and the company name. The CustomerID has a Text data type.

```vb
Sub cmboCompanyName_AfterUpdate() 
    Dim strNewRecord As String 
    strNewRecord = "SELECT * FROM Customers " _ 
        &; " WHERE CustomerID = '" _ 
        &; Me!cmboCompanyName.Value &; "'" 
    Me.RecordSource = strNewRecord 
End Sub
```

The following example shows how to use a Structured Query Language (SQL) statement to establish the data source of a report as it is opened.

**Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)

```vb
Private Sub Report_Open(Cancel As Integer)

    On Error GoTo Error_Handler

    Me.Caption = ?My Application?

    DoCmd.OpenForm FormName:=?frmReportSelector_MemberList?, _
    Windowmode:=acDialog

    ?Cancel the report if ?cancel? was selected on the dialog form.

    If Forms!frmReportSelector_MemberList!txtContinue = ?no? Then
        Cancel = True
        GoTo Exit_Procedure
    End If
    Me.RecordSource = ReplaceWhereClause(Me.RecordSource, _
      Forms!frmReportSelector_MemberList!txtWhereClause)

Exit_Procedure:
    Exit Sub

Error_Handler:
    MsgBox Err.Number &; ?: ? &; Err.Description
    Resume Exit_Procedure
    Resume

End Sub
```


## About the contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also

[Report Object](report-object-access.md)

