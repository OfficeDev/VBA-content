---
title: MailMergeDataSource.TableName Property (Word)
keywords: vbawd10.chm152895505
f1_keywords:
- vbawd10.chm152895505
ms.prod: word
api_name:
- Word.MailMergeDataSource.TableName
ms.assetid: 0dd6f6de-a4b3-383f-d2eb-c76539540d73
ms.date: 06/08/2017
---


# MailMergeDataSource.TableName Property (Word)

Returns a  **String** with the SQL query used to retrieve the records from the data source file attached to a mail merge document. Read-only.


## Syntax

 _expression_ . **TableName**

 _expression_ A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Remarks

This property may be blank if the table name is unknown or not applicable to the current data source.


## Example

This example checks to see if the Customers table is the name of the table in the attached data source. If not, it attaches the Customers table in the Northwind database.


 **Note**  This example uses the Visual Basic  **InStr** function, which returns the position of the first character in the second string if it exists in the first string. A value of zero (0) is returned if the first string does not contain the second string. Setting the conditional value to less than one (1) indicates that the attached table is not named Customers.


```vb
Sub DataSourceTable() 
 With ActiveDocument.MailMerge 
 If InStr(1, .DataSource.TableName, "Customers") < 1 Then 
 .OpenDataSource Name:="C:\ProgramFiles\Microsoft Office\Office\" &; _ 
 "Samples\Northwind.mdb", LinkToSource:=True, _ 
 AddToRecentFiles:=False, Connection:="TABLE Customers" 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)

