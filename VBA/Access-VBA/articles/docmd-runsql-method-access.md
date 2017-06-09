---
title: DoCmd.RunSQL Method (Access)
keywords: vbaac10.chm4176
f1_keywords:
- vbaac10.chm4176
ms.prod: access
api_name:
- Access.DoCmd.RunSQL
ms.assetid: 5d61f75a-b220-cc2c-edea-51a6d4f9f106
ms.date: 06/08/2017
---


# DoCmd.RunSQL Method (Access)

The  **RunSQL** method carries out the RunSQL action in Visual Basic.


## Syntax

 _expression_. **RunSQL**( ** _SQLStatement_**, ** _UseTransaction_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SQLStatement_|Required|**Variant**|A string expression that's a valid SQL statement for an action query or a data-definition query. It uses an INSERT INTO, DELETE, SELECT...INTO, UPDATE, CREATE TABLE, ALTER TABLE, DROP TABLE, CREATE INDEX, or DROP INDEX statement. Include an IN clause if you want to access another database.|
| _UseTransaction_|Optional|**Variant**|Use  **True** (?1) to include this query in a transaction. Use **False** (0) if you don't want to use a transaction. If you leave this argument blank, the default ( **True** ) is assumed.|

## Remarks

You can use the RunSQL action to run a Microsoft Access action query by using the corresponding SQL statement. You can also run a data-definition query.

This method only applies to Microsoft Access databases.

The maximum length of the  _sqlstatement_ argument is 32,768 characters (unlike the SQL Statement action argument in the Macro window, whose maximum length is 256 characters).

 **Links provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Calling Action Queries](http://www.utteraccess.com/wiki/index.php/Calling_Action_Queries)
    
- [Queries/SQL: Use Statements Stored in Table](http://www.utteraccess.com/wiki/index.php/Queries/SQL:_Use_Statements_Stored_in_Table)
    

## Example

The following example updates the Employees table, changing each sales manager's title to Regional Sales Manager:


```vb
Public Sub DoSQL() 
 
    Dim SQL As String 
     
    SQL = "UPDATE Employees " &; _ 
          "SET Employees.Title = 'Regional Sales Manager' " &; _ 
          "WHERE Employees.Title = 'Sales Manager'" 
 
    DoCmd.RunSQL SQL 
     
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[DoCmd Object](docmd-object-access.md)

