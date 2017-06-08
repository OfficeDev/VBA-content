---
title: Apply a Filter When Opening a Form or Report
ms.prod: access
ms.assetid: d7a43e62-3003-d411-2128-dffe0536e119
ms.date: 06/08/2017
---


# Apply a Filter When Opening a Form or Report

When you use Visual Basic for Applications (VBA) code to open a form or report, you may want to specify which records to display. You can specify the records to display in the form or report in several ways. A common approach is to display a custom dialog box in which the user enters criteria for the underlying query of the form or report. To get the criteria, you refer to the controls in the dialog box. The following sections describe three ways you can use criteria entered in a custom dialog box to filter records.


## Using the wherecondition argument

The  _wherecondition_ argument of the **[OpenForm](docmd-openform-method-access.md)** or **[OpenReport](docmd-openreport-method-access.md)** method or action is the simplest way to get criteria in situations where a user is providing only one value. For example, you could display a form that prompts users to select an order ID for the invoice they want to print. If you are using an event procedure, you can apply a filter that displays only one record by adding an argument to the **OpenReport** method, as shown in the following line of code:


```vb
DoCmd.OpenReport "Invoice", acViewPreview, , "OrderID = " &; OrderID 

```

The  `"OrderID = "` in the filter expression refers to the OrderID field in the Invoice report's underlying query. The OrderID on the right side of the expression refers to the value the user selected from the OrderID list in the dialog box. The expression concatenates the two, causing the report to include only the invoice for the record the user selected.

The  _wherecondition_ argument is applied only by the event procedure specified for the **OnClick** event of the button that runs the **OpenForm** or **OpenReport** method. This gives you the flexibility of using any number of different dialog boxes to open the same form or report and applying different sets of criteria depending on what the user wants to do. For example, the user may want to print an invoice for a certain customer or view orders only for a certain product.

You can use the  _wherecondition_ argument to set criteria for more than one field, but if you do, the argument setting quickly becomes long and complicated. In those situations, specifying criteria in a query may be easier.


## Using a Query as a Filter

A separate query, sometimes called a filter query, can refer to the controls on your dialog box to get its criteria. Using this approach, you filter the records in a form or report by setting the  _filtername_ argument of the **OpenForm** or **OpenReport** method or action to the name of the filter query you create. The filter query must include all tables in the record source of the form or report you are opening. Additionally, the filter query must either include all the fields in the form or report you are opening, or you must set its **OutputAllFields** property to **Yes**.

After you create and save the query to use as a filter, set the  _filtername_ argument of the **OpenReport** method or action to the name of the filter query. The _filtername_ argument applies the specified filter query each time the **OpenReport** method runs.

Using a query as a filter to set the criteria has advantages similar to using the  _wherecondition_ argument of the **OpenForm** or **OpenReport** method. A filter query gives you the same flexibility of using more than one dialog box to open the same form or report and applying different sets of criteria depending on what a user wants to do.


## Directly Referring to Dialog Box Controls in the Underlying Query of a Form or Report

You can also refer to the dialog box controls directly in the underlying query of a form or report instead of through the arguments of the  **OpenForm** or **OpenReport** method. Using this approach, the **OpenForm** or **OpenReport** method or action requires no _wherecondition_ or _filtername_ argument. Instead, each time you open a form or report, its underlying query looks for the dialog box to get its criteria. However, if a user opens the form or report in the Database window rather than through your dialog box, Access displays a parameter box that prompts the user for the dialog box value.


