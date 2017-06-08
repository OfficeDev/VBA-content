---
title: Application.DVar Method (Access)
keywords: vbaac10.chm12531
f1_keywords:
- vbaac10.chm12531
ms.prod: access
api_name:
- Access.Application.DVar
ms.assetid: e1566391-4aac-548f-6475-6a8ee63a2bb7
ms.date: 06/08/2017
---


# Application.DVar Method (Access)

Estimates the variance across a sample in a specified set of records (a domain).


## Syntax

 _expression_. **DVar**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|An expression that identifies the numeric field on which you want to find the variance. It can be a string expression identifying a field from a table or query, or it can be an expression that performs a [calculation on data in that field](calculate-fields-in-domain-aggregate-functions.md) . In _expr_, you can include the name field in a table, a control on a form, a constant, or a function. If  _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function. Any field included in _expr_ must be a numeric field.|
| _Domain_|Required|**String**|A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.|
| _Criteria_|Optional|**Variant**|An optional string expression used to restrict the range of data on which the  **DVar** function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the **DVar** function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise the  **DVar** function returns a **Null**.|

### Return Value

Variant


## Remarks

For example, you could use the  **DVar** function to calculate the variance across a set of students' test scores.

If  _domain_ refers to fewer than two records or if fewer than two records satisfy _criteria_, the  **DVar** functions return a **Null**, indicating that a variance can't be calculated.

Whether you use the  **DVar** function in a macro, module, query expression, or calculated control, you must construct the _criteria_ argument carefully to ensure that it will be evaluated correctly.

You can use the  **DVar** function to specify criteria in the **Criteria** row of a select query, in a calculated field expression in a query, or in the **Update To** row of an update query.


 **Note**  You can use the  **DVar** function or the **Var** function in a calculated field expression in a totals query. If you use the **DVar** function, values are calculated before data is grouped. If you use the **Var** function, the data is grouped before values in the field expression are evaluated.

Use the  **DVar** function in a calculated control when you need to specify _criteria_ to restrict the range of data on which the function is performed. For example, to display a variance for orders to be shipped to California, set the **ControlSource** property of a text box to the following expression:




```
=DVar("[Freight]", "Orders", "[ShipRegion] = 'CA'")
```

If you simply want to find the standard deviation across all records in  _domain_, use the  **Var** or **VarP** function.


## Example

The following example returns estimates of the variance for a population and a population sample for orders shipped to the United Kingdom. The domain is an Orders table. The  _criteria_ argument restricts the resulting set of records to those for which ShipCountry equals UK.


```vb
Dim dblX As Double 
Dim dblY As Double 
 
' Sample estimate. 
dblX = DVar("[Freight]", "Orders", "[ShipCountry] = 'UK'") 
 
' Population estimate. 
dblY = DVarP("[Freight]", "Orders", "[ShipCountry] = 'UK'")
```

The next example returns estimates by using a variable,  `strCountry`, in the  _criteria_ argument. Note that single quotation marks (') are included in the string expression, so that when the strings are concatenated, the string literal `UK` will be enclosed in single quotation marks.




```vb
Dim strCountry As String| 
Dim dblX As Double 
 
strCountry = "UK" 
 
dblX = DVar("[Freight]", "Orders", "[ShipCountry] = '" _ 
    &; strCountry &; "'")
```



The following examples show how to use various types of criteria with the  **DVar** function.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```js
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DVar("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DVar("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DVar("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DVar("[FieldName]", "TableName", "[Criteria] = " &; Forms!FormName!ControlName)

    ' Strings
    variable = DVar("[FieldName]", "TableName", "[Criteria] = '" &; Forms!FormName!ControlName &; "'")

    ' Dates
    variable = DVar("[FieldName]", "TableName", "[Criteria] = #" &; Forms!FormName!ControlName &; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DVar("[FieldName]", "TableName", "[Criteria1] = " &; Forms![FormName]![Control1] _
             &; " AND [Criteria2] = '" &; Forms![FormName]![Control2] &; "'" _
            &; " AND [Criteria3] =#" &; Forms![FormName]![Control3] &; "#")
    
    ' Use two fields from a single record.
    variable = DVar("[LastName] &; ', ' &; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DVar("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DVar("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

