
# Shape Append Clause

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection0)
[Description](#sectionSection1)
[Remarks](#sectionSection2)
[Remarks](#sectionSection3)


The shape command APPEND clause appends a column or columns to a  **Recordset**. Often these columns are chapter columns, which refer to a child **Recordset**.

## Syntax
<a name="sectionSection0"> </a>


```
 
SHAPE [parent-command  [[AS] parent-alias ]] APPEND column-list
```


## Description
<a name="sectionSection1"> </a>

The parts of this clause are as follows:


-  _parent-command_
    
- Zero or one of the following (you may omit the  _parent-command_ entirely):
    
      - A provider command within curly braces ("{}") that returns a  **Recordset** object. The command is issued to the underlying data provider, and its syntax depends on the requirements of that provider. This will typically be the SQL language, although ADO does not require any particular query language.
    
  - Another shape command embedded in parentheses.
    
  - The TABLE keyword, followed by the name of a table in the data provider.
    
-  _parent-alias_
    
- An optional alias that refers to the parent  **Recordset**.
    
-  _column-list_
    
- One or more of the following:
    
      - An aggregate column.
    
  - A calculated column.
    
  - A new column created with the NEW clause.
    
  - A chapter column. A chapter column definition is enclosed in parentheses ("()"). See syntax below:
    

```sql
 
SHAPE [parent-command  [[AS] parent-alias ]] 
 APPEND (child-recordset  [ [[AS] child-alias ] 
 RELATE parent-column  TO child-column  | PARAMETER param-number , ... ]) 
 [[AS] chapter-alias ] 
 [, ... ] 

```


-  _child-recordset_
    
- 
    
      - A provider command within curly braces ("{}") that returns a  **Recordset** object. The command is issued to the underlying data provider, and its syntax depends on the requirements of that provider. This will typically be the SQL language, although ADO does not require any particular query language.
    
  - Another shape command embedded in parentheses.
    
  - The name of an existing shaped  **Recordset**.
    
  - The TABLE keyword, followed by the name of a table in the data provider.
    
-  _child-alias_
    
- An alias that refers to the child  **Recordset**.
    
-  _parent-column_
    
- A column in the  **Recordset** returned by the _parent-command._
    
-  _child-column_
    
- A column in the  **Recordset** returned by the _child-command_.
    
-  _param-number_
    
- See [Operation of Parameterized Commands](71edbd16-21db-7afa-356b-d8e7afb92b3a.md).
    
-  _chapter-alias_
    
- An alias that refers to the chapter column appended to the parent.
    

 **Note**  The  _"parent-column_ TO _child-column"_ clause is actually a list, where each relation defined is separated by a comma.


 **Note**  The clause after the APPEND keyword is actually a list, where each clause is separated by a comma and defines another column to be appended to the parent.


## Remarks
<a name="sectionSection2"> </a>

When you construct provider commands from user input as part of a SHAPE command, SHAPE will treat the user-supplied a provider command as an opaque string and pass them faithfully to the provider. For example, in the following SHAPE command,


```
 
SHAPE {select * from t1} APPEND ({select * from t2} RELATE k1 TO k2) 

```

SHAPE will execute two commands:  `select * from t1` and ( `select * from t2 RELATE k1 TO k2)`. If the user supplies a compound command consisting of multiple provider commands separated by semicolons, SHAPE is not able to discern the difference. So in the following SHAPE command,




```
 
SHAPE {select * from t1; drop table t1} APPEND ({select * from t2} RELATE k1 TO k2) 

```

SHAPE executes  `select * from t1; drop table t1` and ( `select * from t2 RELATE k1 TO k2),` not realizing that `drop table t1` is a separate and in this case, dangerous, provider command. Applications must always validate the user input to prevent such potential hacker attacks from happening.


## Remarks
<a name="sectionSection3"> </a>

When you construct provider commands from user input as part of a SHAPE command, SHAPE will treat the user-supplied a provider command as an opaque string and pass them faithfully to the provider. For example, in the following SHAPE command,


```
 
SHAPE {select * from t1} APPEND ({select * from t2} RELATE k1 TO k2) 

```

SHAPE will execute two commands:  `select * from t1` and ( `select * from t2 RELATE k1 TO k2)`. If the user supplies a compound command consisting of multiple provider commands separated by semicolons, SHAPE is not able to discern the difference. So in the following SHAPE command,




```
 
SHAPE {select * from t1; drop table t1} APPEND ({select * from t2} RELATE k1 TO k2) 

```

SHAPE executes  `select * from t1; drop table t1` and ( `select * from t2 RELATE k1 TO k2),` not realizing that `drop table t1` is a separate and in this case, dangerous, provider command. Applications must always validate the user input to prevent such potential hacker attacks from happening.

