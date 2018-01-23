---
title: Report.NoData Event (Access)
keywords: vbaac10.chm13881
f1_keywords:
- vbaac10.chm13881
ms.prod: access
api_name:
- Access.Report.NoData
ms.assetid: fa5f22b1-3695-bd16-2ca3-b2a1cc1f1d94
ms.date: 11/30/2017
---


# Report.NoData Event (Access)

The **NoData** event occurs after Microsoft Access formats a report for printing that has no data (the report is bound to an empty recordset), but before the report is printed. You can use this event to cancel printing of a blank report.

## Syntax

_expression_. **NoData**(**_Cancel_**)

_expression_ A variable that represents a **Report** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines whether to print the report. Setting the _Cancel_ argument to **True** (?1) prevents the report from printing. You can also use the **CancelEvent** method of the **DoCmd** object to cancel printing the report.|

## Remarks

To run a macro or event procedure when this event occurs, set the **[OnNoData](report-onnodata-property-access.md)** property to the name of the macro or to [Event Procedure].

If the report isn't bound to a table or query (by using the report's **[RecordSource](report-recordsource-property-access.md)** property), the **NoData** event doesn't occur.

This event occurs after the  **Format** events for the report, but before the first **Print** event.

This event doesn't occur for subreports. If you want to hide controls on a subreport when the subreport has no data, so that the controls don't print in this case, you can use the **HasData** property in a macro or event procedure that runs when the **Format** or **Print** event occurs.

The **NoData** event occurs before the first **Page** event for the report.


## Example

The following example shows how to cancel printing a report when it has no data. A message box notifying the user that the printing has been canceled is also displayed. 

To try this example, add the following event procedure to a report. Try running the report when it contains no data. 

```vb
Private Sub Report_NoData(Cancel As Integer) 
    MsgBox "The report has no data." &; _ 
         chr(13) &; "Printing is canceled. " &; _ 
         chr(13) &; "Check the data source for the " &; _ 
         chr(13) &; "report. Make sure you entered " &; _ 
         chr(13) &; "the correct criteria (for " &; _ 
         chr(13) &; "example, a valid range of " &; _ 
         chr(13) &; "dates),." vbOKOnly + vbInformation 
    Cancel = True 
End Sub 
```


The following example shows how to use the **NoData** event of a report to prevent the report form opening when there is no data to be displayed.


**Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)

```vb
Private Sub Report_NoData(Cancel As Integer)

    'Add code here that will be executed if no data
    'was returned by the Report's RecordSource
    MsgBox "No customers ordered this product this month. " &; _
        "The report will now close."
    Cancel = True

End Sub
```


## About the contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also

[Report Object](report-object-access.md)

