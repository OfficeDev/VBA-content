---
title: Application.DSum Method (Access)
keywords: vbaac10.chm12527
f1_keywords:
- vbaac10.chm12527
ms.prod: access
api_name:
- Access.Application.DSum
ms.assetid: 53a3cfd4-a5e3-d0c5-1727-070c99d2b984
ms.date: 06/08/2017
---


# Application.DSum Method (Access)

You can use the  **DSum** function to calculate the sum of a set of values in a specified set of records (a domain). .


## Syntax

 _expression_. **DSum**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|An expression that identifies the numeric field whose values you want to total. It can be a string expression identifying a field in a table or query, or it can be an expression that performs a [calculation on data in that field](calculate-fields-in-domain-aggregate-functions.md) . In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If  _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.|
| _Domain_|Required|**String**|A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.|
| _Criteria_|Optional|**Variant**|An optional string expression used to restrict the range of data on which the  **DSum** function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the **DSum** function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the  **DSum** function returns a **Null**.|

### Return Value

Variant


## Remarks

For example, you could use the  **DSum** function in a calculated field expression in a query to calculate the total sales made by a particular employee over a period of time. Or you could use the **DSum** function in a calculated control to display a running sum of sales for a particular product.

If no record satisfies the  _criteria_ argument or if domain contains no records, the **DSum** function returns a **Null**.

Whether you use the  **DSum** function in a macro, module, query expression, or calculated control, you must construct the _criteria_ argument carefully to ensure that it will be evaluated correctly.

You can use the  **DSum** function to specify criteria in the Criteria row of a query, in a calculated field in a query expression, or in the Update To row of an update query.


 **Note**  You can use either the  **DSum** or **Sum** function in a calculated field expression in a totals query. If you use the **DSum** function, values are calculated before data is grouped. If you use the **Sum** function, the data is grouped before values in the field expression are evaluated.

You may want to use the  **DSum** function when you need to display the sum of a set of values from a field that is not in the record source for your form or report. For example, suppose you have a form that displays information about a particular product. You could use the **DSum** function to maintain a running total of sales of that product in a calculated control.

If you need to maintain a running total in a control on a report, you can use the  **RunningSum** property of that control if the field on which it is based is included in the record source for the report. Use the **DSum** function to maintain a running sum on a form.


## Example

The following example totals the values from the Freight field for orders shipped to the United Kingdom. The domain is an Orders table. The  _criteria_ argument restricts the resulting set of records to those for which ShipCountry equals UK.


```vb
Dim curX As Currency 
curX = DSum("[Freight]", "Orders", "[ShipCountry] = 'UK'")
```

The next example calculates a total by using two separate criteria. Note that single quotation marks (') and number signs (#) are included in the string expression, so that when the strings are concatenated, the string literal will be enclosed in single quotation marks, and the date will be enclosed in number signs.




```vb
Dim curX As Currency 
curX = DSum("[Freight]", "Orders", _ 
    "[ShipCountry] = 'UK' AND [ShippedDate] > #1-1-95#")
```



The following examples show how to use various types of criteria with the  **DSum** function.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```js
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DSum("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DSum("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DSum("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DSum("[FieldName]", "TableName", "[Criteria] = " &; Forms!FormName!ControlName)

    ' Strings
    variable = DSum("[FieldName]", "TableName", "[Criteria] = '" &; Forms!FormName!ControlName &; "'")

    ' Dates
    variable = DSum("[FieldName]", "TableName", "[Criteria] = #" &; Forms!FormName!ControlName &; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DSum("[FieldName]", "TableName", "[Criteria1] = " &; Forms![FormName]![Control1] _
             &; " AND [Criteria2] = '" &; Forms![FormName]![Control2] &; "'" _
            &; " AND [Criteria3] =#" &; Forms![FormName]![Control3] &; "#")
    
    ' Use two fields from a single record.
    variable = DSum("[LastName] &; ', ' &; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DSum("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DSum("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

