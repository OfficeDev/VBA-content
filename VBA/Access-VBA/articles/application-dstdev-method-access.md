---
title: Application.DStDev Method (Access)
keywords: vbaac10.chm12533
f1_keywords:
- vbaac10.chm12533
ms.prod: access
api_name:
- Access.Application.DStDev
ms.assetid: 401b4e16-dfd4-7256-b03d-f3915c5f9ca5
ms.date: 06/08/2017
---


# Application.DStDev Method (Access)

Estimates the standard deviation across a population sample in a specified set of records (a domain). .


## Syntax

 _expression_. **DStDev**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|An expression that identifies the numeric field on which you want to find the standard deviation. It can be a string expression identifying a field from a table or query, or it can be an expression that performs a [calculation on data in that field](calculate-fields-in-domain-aggregate-functions.md). In  _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If  _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.|
| _Domain_|Required|**String**|A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.|
| _Criteria_|Optional|**Variant**|An optional string expression used to restrict the range of data on which the  **DStDev** function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the **DStDev** function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the  **DStDev** function will return a **Null**.|

### Return Value

Variant


## Remarks

For example, you could use the  **DStDev** function in a module to calculate the standard deviation across a set of students' test scores.

If  _domain_ refers to fewer than two records or if fewer than two records satisfy _criteria_, the  **DStDev** function returns a **Null**, indicating that a standard deviation can't be calculated.

You can use the  **DStDev** function to specify criteria in the Criteria row of a select query. For example, you could create a query on an Orders table and a Products table to display all products for which the freight cost fell above the mean plus the standard deviation for freight cost. The Criteria row beneath the Freight field would contain the following expression:




```
>(DStDev("[Freight]", "Orders") + DAvg("[Freight]", "Orders"))
```

You can use the  **DStDev** function in a calculated field expression of a query, or in the Update To row of an update query.


 **Note**  You can use the  **DStDev** and **DStDevP** functions or the **StDev** and **StDevP** functions in a calculated field expression of a totals query. If you use the **DStDev** or **DStDevP** function, values are calculated before data is grouped. If you use the **StDev** or **StDevP** function, the data is grouped before values in the field expression are evaluated.

Use the  **DStDev** function in a calculated control when you need to specify criteria to restrict the range of data on which the function is performed. For example, to display standard deviation for orders to be shipped to California, set the **ControlSource** property of a text box to the following expression:




```
=DStDev("[Freight]", "Orders", "[ShipRegion] = 'CA'")
```

If you simply want to find the standard deviation across all records in  _domain_, use the  **StDev** or **StDevP** function.

If the data type of the field from which  _expr_ is derived is a number, the **DStDev** function returns a **Double** data type. If you use the **DStDev** function in a calculated control, include a data type conversion function in the expression to improve performance.


## Example

The following example returns estimates of the standard deviation for a population and a population sample for orders shipped to the United Kingdom. The domain is an Orders table. The  _criteria_ argument restricts the resulting set of records to those for which the ShipCountry value is UK.


```vb
Dim dblX As Double 
Dim dblY As Double 
 
' Sample estimate. 
dblX = DStDev("[Freight]", "Orders", "[ShipCountry] = 'UK'") 
 
' Population estimate. 
dblY = DStDevP("[Freight]", "Orders", "[ShipCountry] = 'UK'")
```

The next example calculates the same estimates by using a variable,  `strCountry`, in the  _criteria_ argument. Note that single quotation marks (') are included in the string expression, so that when the strings are concatenated, the string literal `UK` will be enclosed in single quotation marks.




```vb
Dim strCountry As String 
Dim dblX As Double 
Dim dblY As Double 
 
strCountry = "UK" 
 
dblX = DStDev("[Freight]", "Orders", _ 
    "[ShipCountry] = '" &; strCountry &; "'") 
 
dblY = DStDevP("[Freight]", "Orders", _ 
    "[ShipCountry] = '" &; strCountry &; "'")
```



The following examples show how to use various types of criteria with the  **DStDev** function.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```js
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DStDev("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DStDev("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DStDev("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DStDev("[FieldName]", "TableName", "[Criteria] = " &; Forms!FormName!ControlName)

    ' Strings
    variable = DStDev("[FieldName]", "TableName", "[Criteria] = '" &; Forms!FormName!ControlName &; "'")

    ' Dates
    variable = DStDev("[FieldName]", "TableName", "[Criteria] = #" &; Forms!FormName!ControlName &; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DStDev("[FieldName]", "TableName", "[Criteria1] = " &; Forms![FormName]![Control1] _
             &; " AND [Criteria2] = '" &; Forms![FormName]![Control2] &; "'" _
            &; " AND [Criteria3] =#" &; Forms![FormName]![Control3] &; "#")
    
    ' Use two fields from a single record.
    variable = DStDev("[LastName] &; ', ' &; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DStDev("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DStDev("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

