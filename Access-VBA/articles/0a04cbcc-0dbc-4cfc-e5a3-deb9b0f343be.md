
# Application.DLast Method (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)
 [About the Contributors](#AboutContributors)


You can use the  **DLast** function to return a random record from a particular field in a table or query when you simply need any value from that field. .


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **DLast**( **_Expr_**,  **_Domain_**,  **_Criteria_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Expr|Required| **String**|An expression that identifies the field from which you want to find the first or last value. It can be either a string expression identifying a field in a table or query, or an expression that performs a  [calculation on data in that field](73c27d1c-0a3c-03e4-c17c-337133d7b316.md). In expr, you can include the name of a field in a table, a control on a form, a constant, or a function. If expr includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.|
|Domain|Required| **String**|A string expression identifying the set of records that constitutes the domain.|
|Criteria|Optional| **Variant**|An optional string expression used to restrict the range of data on which the  **DLast** function is performed. For example,criteria is often equivalent to the WHERE clause in an SQL expression, without the wrd WHERE. Ifcriteria is omitted, the **DLast** function evaluatesexpr against the entire domain. Any field that is included incriteria must also be a field indomain; otherwise, the  **DLast** function returns a **Null**.|

### Return Value

Variant


## Remarks
<a name="sectionSection1"> </a>




 **Note**   If you want to return the first or last record in a set of records (a domain), you should create a query sorted as either ascending or descending and set the **TopValues** property to 1. From Visual Basic, you can also create an ADO **Recordset** object and use the **MoveFirst** or **MoveLast** method to return the first or last record in a set of records.


## Example
<a name="sectionSection2"> </a>



The following examples show how to use various types of criteria with the  **DLast** function.

 **Sample code provided by:**
![Community Member Icon](../images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The [UtterAccess](http://www.utteraccess.com) community




```
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DLast("[FieldName]", "TableName", "[Criteria] = " &amp; Forms!FormName!ControlName)

    ' Strings
    variable = DLast("[FieldName]", "TableName", "[Criteria] = '" &amp; Forms!FormName!ControlName &amp; "'")

    ' Dates
    variable = DLast("[FieldName]", "TableName", "[Criteria] = #" &amp; Forms!FormName!ControlName &amp; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DLast("[FieldName]", "TableName", "[Criteria1] = " &amp; Forms![FormName]![Control1] _
             &amp; " AND [Criteria2] = '" &amp; Forms![FormName]![Control2] &amp; "'" _
            &amp; " AND [Criteria3] =#" &amp; Forms![FormName]![Control3] &amp; "#")
    
    ' Use two fields from a single record.
    variable = DLast("[LastName] &amp; ', ' &amp; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DLast("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DLast("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


 [Application Object](aefb0713-97e6-e2c7-e530-8fd2e1316a55.md)
#### Other resources


 [Application Object Members](3ab5276c-d52a-72a9-244c-ec92ead48811.md)
