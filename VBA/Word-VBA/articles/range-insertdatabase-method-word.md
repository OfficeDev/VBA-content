---
title: Range.InsertDatabase Method (Word)
keywords: vbawd10.chm157155522
f1_keywords:
- vbawd10.chm157155522
ms.prod: word
api_name:
- Word.Range.InsertDatabase
ms.assetid: c8bcddda-0943-9619-e5ee-9ef0956b0f43
ms.date: 06/08/2017
---


# Range.InsertDatabase Method (Word)

Retrieves data from a data source (for example, a separate Microsoft Word document, a Microsoft Excel worksheet, or a Microsoft Access database) and inserts the data as a table in place of the specified range.


## Syntax

 _expression_ . **InsertDatabase**( **_Format_** , **_Style_** , **_LinkToSource_** , **_Connection_** , **_SQLStatement_** , **_SQLStatement1_** , **_PasswordDocument_** , **_PasswordTemplate_** , **_WritePasswordDocument_** , **_WritePasswordTemplate_** , **_DataSource_** , **_From_** , **_To_** , **_IncludeFields_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Format_|Optional| **Variant**|A format listed in the  **Formats** box in the **Table AutoFormat** dialog box ( **Table** menu). Can be any of the **WdTableFormat** constants. A border is applied to the cells in the table by default.|
| _Style_|Optional| **Variant**|The attributes of the AutoFormat specified by Format that are applied to the table.|
| _LinkToSource_|Optional| **Variant**| **True** to establish a link between the new table and the data source.|
| _Connection_|Optional| **Variant**|A range within which to perform the query specified by SQLStatement.|
| _SQLStatement_|Optional| **String**|An optional query string that retrieves a subset of the data in a primary data source to be inserted into the document.|
| _SQLStatement1_|Optional| **String**|If the query string is longer than 255 characters, SQLStatement denotes the first portion of the string and SQLStatement1 denotes the second portion.|
| _PasswordDocument_|Optional| **Variant**|The password (if any) required to open the data source. (See Remarks below.)|
| _PasswordTemplate_|Optional| **Variant**|If the data source is a Word document, this argument is the password (if any) required to open the attached template. (See Remarks below.)|
| _WritePasswordDocument_|Optional| **Variant**|The password required to save changes to the document. (See Remarks below.)|
| _WritePasswordTemplate_|Optional| **Variant**|The password required to save changes to the template. (See Remarks below.)|
| _DataSource_|Optional| **Variant**|The path and file name of the data source.|
| _From_|Optional| **Variant**|The number of the first record in the range of records to be inserted.|
| _To_|Optional| **Variant**|The number of the last record in the range of records to be inserted.|
| _IncludeFields_|Optional| **Variant**| **True** to include field names from the data source in the first row of the new table.|

## Security

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


 **Security Note**  




 **Security Note**  




## Remarks

The value of the Style argument can be the sum of any combination of the following values:



|**Value**|**Meaning**|
|:-----|:-----|
|0 (zero)|None|
|1|Borders|
|2|Shading|
|4|Font|
|8|Color|
|16|Auto Fit|
|32|Heading Rows|
|64|Last Row|
|128|First Column|
|256|Last Column|
How you specify the Connection argument depends on how data is retrieved. For example:


- When retrieving data through Open Database Connectivity (ODBC), you specify a connection string.
    
- When retrieving data from Excel by using dynamic data exchange (DDE), you specify a named range or "Entire Spreadsheet."
 **Security Note**  


    
- When retrieving data from Access, you specify the word "Table" or "Query" followed by the name of a table or query.
    

## Example

This example inserts an Excel spreadsheet named "Data.xls" after the selection. The Style value (191) is a combination of the numbers 1, 2, 4, 8, 16, 32, and 128.


```vb
With Selection 
    .Collapse Direction:=wdCollapseEnd 
    .Range.InsertDatabase _ 
        Format:=wdTableFormatSimple2, Style:=191, _ 
        LinkToSource:=False, Connection:="Entire Spreadsheet", _ 
        DataSource:="C:\MSOffice\Excel\Data.xls" 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

