---
title: Application.DMin Method (Access)
keywords: vbaac10.chm12525
f1_keywords:
- vbaac10.chm12525
ms.prod: access
api_name:
- Access.Application.DMin
ms.assetid: d41b1852-7d97-ddfe-d071-8a1a7b42359b
ms.date: 06/08/2017
---


# Application.DMin Method (Access)

You can use  **DMin** function to determine minnimum value in a specified set of records (a domain). .


## Syntax

 _expression_. **DMin**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|An expression that identifies the field for which you want to find the minimum or maximum value. It can be a string expression identifying a field in a table or query, or it can be an expression that performs [calculation on data in that field](calculate-fields-in-domain-aggregate-functions.md) . In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If  _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.|
| _Domain_|Required|**String**|A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.|
| _Criteria_|Optional|**Variant**|An optional string expression used to restrict the range of data on which the  **DMin** function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the **DMin** function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_, otherwise the  **DMin** function returns a **Null**.|

### Return Value

Variant


## Remarks

For example, you could use the  **DMin** function in calculated controls on a report to display smallest order amount for a particular customer. Or you could use the **DMin** function in a query expression to display all orders with a discount greater than the minimum possible discount.

The  **DMin** function returns the minimum values that satisfy _criteria_. If  _expr_ identifies numeric data, the **DMin** function returns numeric values. If _expr_ identifies string data, they return the string that is first or last alphabetically.

The  **DMin** function ignores **Null** values in the field referenced by _expr_. However, if no record satisfies  _criteria_ or if _domain_ contains no records, the **DMin** function returns a **Null**.

You can use the  **DMin** function to specify criteria in the Criteria row of a query, in a calculated field expression in a query, or in the Update To row of an update query.


 **Note**  You can use the  **DMin** function or the **Min** function in a calculated field expression of a totals query. If you use the **DMin** function, values are evaluated before the data is grouped. If you use the **Min** function, the data is grouped before values in the field expression are evaluated.

If you simply want to find the minimum or maximum value of all records in  _domain_, use the  **Min** or **Max** function.

Although you can use the  **DMin** or **DMax** function to find the minimum or maximum value from a field in a foreign table, it may be more efficient to create a query that contains the fields that you need from both tables and base your form or report on that query.


## Example

The following example returns the lowest and highest values from the Freight field for orders shipped to the United Kingdom. The domain is an Orders table. The  _criteria_ argument restricts the resulting set of records to those for which ShipCountry equals UK.


```vb
Dim curX As Currency 
Dim curY As Currency 
 
curX = DMin("[Freight]", "Orders", "[ShipCountry] = 'UK'") 
curY = DMax("[Freight]", "Orders", "[ShipCountry] = 'UK'")
```

In the next example, the criteria expression includes a variable,  `dteOrderDate`. Note that number signs (#) are included in the string expression, so that when the strings are concatenated, they will enclose the date.




```vb
Dim dteOrderDate As Date 
Dim curX As Currency 
 
dteOrderDate = #03/30/2000# 
curX = DMin("[Freight]", "Orders", _ 
    "[OrderDate] = #" &; dteOrderDate &; "#")
```



The following examples show how to use various types of criteria with the  **DMin** function.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```js
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DMin("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DMin("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DMin("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DMin("[FieldName]", "TableName", "[Criteria] = " &; Forms!FormName!ControlName)

    ' Strings
    variable = DMin("[FieldName]", "TableName", "[Criteria] = '" &; Forms!FormName!ControlName &; "'")

    ' Dates
    variable = DMin("[FieldName]", "TableName", "[Criteria] = #" &; Forms!FormName!ControlName &; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DMin("[FieldName]", "TableName", "[Criteria1] = " &; Forms![FormName]![Control1] _
             &; " AND [Criteria2] = '" &; Forms![FormName]![Control2] &; "'" _
            &; " AND [Criteria3] =#" &; Forms![FormName]![Control3] &; "#")
    
    ' Use two fields from a single record.
    variable = DMin("[LastName] &; ', ' &; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DMin("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DMin("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

