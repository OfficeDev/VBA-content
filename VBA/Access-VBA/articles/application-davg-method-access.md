---
title: Application.DAvg Method (Access)
keywords: vbaac10.chm12528
f1_keywords:
- vbaac10.chm12528
ms.prod: access
ms.assetid: 966cd884-8693-d1d2-b35b-567e71b7e56d
ms.date: 06/08/2017
---


# Application.DAvg Method (Access)

You can use the  **DAvg** function to calculate the average of a set of values in a specified set of records (a domain).


## Syntax

 _expression_. **DAvg**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|An expression that identifies the field containing the numeric data you want to average. It can be a string expression identifying a field in a table or query, or it can be an expression that performs a calculation on data in that field. In  _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If  _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.|
| _Domain_|Required|**String**|A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.|
| _Criteria_|Optional|**Variant**|An optional string expression used to restrict the range of data on which the  **DAvg** function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the **DAvg** function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise the  **DAvg** function returns a **Null**.|

### Return Value

Variant


## Remarks

For example, you could use the  **DAvg** function in the criteria row of a select query on freight cost to restrict the results to those records where the freight cost exceeds the average. Or you could use an expression including the **DAvg** function in a calculated control and display the average value of previous orders next to the value of a new order.

Records containing  **Null** values aren't included in the calculation of the average.

Whether you use the  **DAvg** function in a macro or module, in a query expression, or in a calculated control, you must construct the _criteria_ argument carefully to ensure that it will be evaluated correctly.

You can use the  **DAvg** function to specify criteria in the Criteria row of a query. For example, suppose you want to view a list of all products ordered in quantities above the average order quantity. You could create a query on the Orders, Order Details, and Products tables, and include the Product Name field and the Quantity field, with the following expression in the Criteria row beneath the Quantity field:




```
>DAvg("[Quantity]", "Orders")
```

You can also use the  **DAvg** function within a calculated field expression in a query, or in the Update To row of an update query.


 **Note**  You can use either the  **DAvg** or **Avg** function in a calculated field expression in a totals query. If you use the **DAvg** function, values are averaged before the data is grouped. If you use the **Avg** function, the data is grouped before values in the field expression are averaged.

Use the  **DAvg** function in a calculated control when you need to specify criteria to restrict the range of data on which the **DAvg** function is performed. For example, to display the average cost of freight for shipments sent to California, set the **ControlSource** property of a text box to the following expression:




```
=DAvg("[Freight]", "Orders", "[ShipRegion] = 'CA'")
```

If you simply want to average all records in  _domain_, use the  **Avg** function.

You can use the  **DAvg** function in a module or macro or in a calculated control on a form if a field that you need to display isn't in the record source on which your form is based. For example, suppose you have a form based on the Orders table, and you want to include the Quantity field from the Order Details table in order to display the average number of items ordered by a particular customer. You can use the **DAvg** function to perform this calculation and display the data on your form.

 **Tips**


- If you use the  **DAvg** function in a calculated control, you may want to place the control on the form header or footer so that the value for this control is not recalculated each time you move to a new record.
    
- If the data type of the field from which  _expr_ is derived is a number, the **DAvg** function returns a **Double** data type. If you use the **DAvg** function in a calculated control, include a data type conversion function in the expression to improve performance.
    
- Although you can use the  **DAvg** function to determine the average of values in a field in a foreign table, it may be more efficient to create a query that contains all of the fields that you need and then base your form or report on that query.
    

## Example

The following function returns the average freight cost for orders shipped on or after a given date. The domain is an Orders table. The  _criteria_ argument restricts the resulting set of records based on the given country and ship date. Note that the keyword **AND** is included in the string to separate the multiple fields in the _criteria_ argument. All records included in the **DAvg** function calculation will have both of these criteria.


```vb
Public Function AvgFreightCost(ByVal strCountry As String, _ 
                               ByVal dteShipDate As Date) As Double 
 
    AvgFreightCost = DAvg("[Freight]", "Orders", _ 
                     "[ShipCountry] = '" &; strCountry &; _ 
                     "'AND [ShippedDate] >= #" &; dteShipDate &; "#") 
 
End Function
```



The following examples show how to use various types of criteria with the  **DAvg** function.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```js
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DAvg("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DAvg("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DAvg("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DAvg("[FieldName]", "TableName", "[Criteria] = " &; Forms!FormName!ControlName)

    ' Strings
    variable = DAvg("[FieldName]", "TableName", "[Criteria] = '" &; Forms!FormName!ControlName &; "'")

    ' Dates
    variable = DAvg("[FieldName]", "TableName", "[Criteria] = #" &; Forms!FormName!ControlName &; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DAvg("[FieldName]", "TableName", "[Criteria1] = " &; Forms![FormName]![Control1] _
             &; " AND [Criteria2] = '" &; Forms![FormName]![Control2] &; "'" _
            &; " AND [Criteria3] =#" &; Forms![FormName]![Control3] &; "#")
    
    ' Use two fields from a single record.
    variable = DAvg("[LastName] &; ', ' &; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DAvg("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DAvg("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

