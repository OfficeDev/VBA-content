---
title: DoCmd.ApplyFilter Method (Access)
keywords: vbaac10.chm4142
f1_keywords:
- vbaac10.chm4142
ms.prod: access
api_name:
- Access.DoCmd.ApplyFilter
ms.assetid: 926c7135-131b-1a7c-465b-a9b2ed71cd7b
ms.date: 06/08/2017
---


# DoCmd.ApplyFilter Method (Access)

The  **ApplyFilter** method carries out the **ApplyFilter** action in Visual Basic.


## Syntax

 _expression_. **ApplyFilter**( ** _FilterName_**, ** _WhereCondition_**, ** _ControlName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FilterName_|Optional|**Variant**|A string expression that is the valid name of a filter or query in the current database. When using this method to apply a server filter, the  _FilterName_ argument must be blank.|
| _WhereCondition_|Optional|**Variant**|A string expression that is a valid SQL WHERE clause without the word WHERE.|
| _ControlName_|Optional|**Variant**||

## Remarks

You can use the ApplyFilter action to apply a filter, a query, or an SQL WHERE clause to a table, form, or report to restrict or sort the records in the table, or the records from the underlying table or query of the form or report. For reports, you can use this action only in a macro specified by the report's  **OnOpen** event property.

You can use this action to apply an SQL WHERE clause only when applying a server filter. A server filter cannot be applied to a stored procedure's record source.


 **Note**  You can use the Filter Name argument if you've already defined a filter that provides the appropriate data. You can use the Where Condition argument to enter the restriction criteria directly. If you use both arguments, Microsoft Access applies the WHERE clause to the results of the filter. You must use one or both arguments.

You can apply a filter or query to a form in Form view or Datasheet view.

The filter and WHERE condition you apply become the setting of the form's  **[Filter](form-filter-property-access.md)** property or the report's **[ServerFilter](report-serverfilter-property-access.md)** property.

When you save a table or form, Access saves any filter currently defined in that object, but will not apply the filter automatically the next time the object is opened (although it will automatically apply any sort you applied to the object before it was saved). If you want to apply a filter automatically when a form is first opened, specify a macro containing the ApplyFilter action or an event procedure containing the  **ApplyFilter** method of the **DoCmd** object as the **OnOpen** event property setting of the form. You can also apply a filter by using the OpenForm or OpenReport action, or their corresponding methods. To apply a filter automatically when a table is first opened, you can open the table by using a macro containing the OpenTable action, followed immediately by the ApplyFilter action.

You must include at least one of the two  **ApplyFilter** method arguments. If you enter a value for both arguments, the _WhereCondition_ argument is applied to the filter.

The maximum length of the  _WhereCondition_ argument is 32,768 characters (unlike the Where Condition action argument in the Macro window, whose maximum length is 256 characters).


## Example

The following example uses the  **ApplyFilter** method to display only records that contain the name "King" in the LastName field:


```vb
DoCmd.ApplyFilter , "LastName = 'King'"
```



The following example shows how to use the  **ApplyFilter** property to filter the records displayed when a toggle button named tglFilter is clicked.

 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```vb
Private Sub tglFilter_Click()
    With tglFilter
        If .Value = True Then
            .Caption = "P/T"
            .StatusBarText = "only full-timers"
            DoCmd.ApplyFilter , "[Hours]=40"
        ElseIf .Value = False Then
            .Caption = "All"
            .StatusBarText = "only part-timers"
            DoCmd.ApplyFilter , "[Hours]<40"
        Else
            .Caption = "F/T"
            .StatusBarText = "all employees"
            DoCmd.ShowAllRecords
            .SetFocus 'to activate the button's statusbar-text
        End If
    End With
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[DoCmd Object](docmd-object-access.md)

