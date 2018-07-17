---
title: Application.CreateForm Method (Access)
keywords: vbaac10.chm12516
f1_keywords:
- vbaac10.chm12516
ms.prod: access
api_name:
- Access.Application.CreateForm
ms.assetid: 113c8f7f-baf1-bf5c-85ce-6dc1f3d3e942
ms.date: 06/08/2017
---


# Application.CreateForm Method (Access)

The  **CreateForm** method creates a form and returns a **[Form](form-object-access.md)** object.


## Syntax

 _expression_. **CreateForm**( ** _Database_**, ** _FormTemplate_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Database_|Optional|**Variant**|The name of the database that contains the form template you want to use to create a form. If you want the current database, omit this argument. If you want to use an open library database, specify the library database with this argument.|
| _FormTemplate_|Optional|**Variant**|The name of the form you want to use as a template to create a new form.|

### Return Value

Form


## Remarks

You can use the  **CreateForm** method when designing a wizard that creates a new form.

The  **CreateForm** method opens a new, minimized form in form Design view.

If the name you use for the  _formtemplate_ argument isn't valid, Visual Basic uses the form template specified by the **Form Template** setting on the **Forms/Reports** tab of the **Options** dialog box.


## Example

This example creates a new form in the Northwind sample database based on the Customers form, and sets its  **RecordSource** property to the Customers table. Run this code from the Northwind sample database.


```vb
Sub NewForm() 
 Dim frm As Form 
 
 ' Create form based on Customers form. 
 Set frm = CreateForm( , "Customers") 
 DoCmd.Restore 
 ' Set RecordSource property to Customers table. 
 frm.RecordSource = "Customers" 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

