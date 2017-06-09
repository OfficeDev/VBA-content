---
title: The Recordset Object Open Method
ms.prod: access
ms.assetid: 5df72473-725c-39f5-a2d0-71466fba70df
ms.date: 06/08/2017
---


# The Recordset Object Open Method

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Source and Options Arguments](#sectionSection0)
[ActiveConnection Argument](#sectionSection1)
[CursorType Argument](#sectionSection2)
[LockType Argument](#sectionSection3)
[Retrieving Multiple Recordsets](#sectionSection4)


Everything you need to open an ADO  **Recordset** is built into the **Open** method. You can use it without explicitly creating any other objects. The syntax of this method is as follows:



```
recordset .OpenSource, ActiveConnection, CursorType, LockType, Options
```

All arguments are optional because the information they pass can be communicated to ADO in other ways. However, understanding each argument will help you to understand many important ADO concepts. The following topics will examine each argument of this method in more detail.

## Source and Options Arguments
<a name="sectionSection0"> </a>

The  _Source_ and _Options_ arguments appear in the same topic because they are closely related.


```
recordset .Open Source, ActiveConnection, CursorType, LockType, Options
```

The  _Source_ argument is a **Variant** that evaluates to a valid **Command** object, a text command (e.g., a SQL statement), a table name, a stored procedure call, a URL, or the name of a file or **Stream** object containing a persistently stored **Recordset**. If _Source_ is a file path name, it can be a full path ("C:\dir\file.rst"), a relative path ("..\file.rst"), or a URL ("http://files/file.rst"). You can also specify this information in the **Recordset** object **Source** property and leave the _Source_ argument blank.

The  _Options_ argument is a **Long** value that indicates either or both of the following:


- How the provider should evaluate the  _Source_ argument if it represents something other than a **Command** object.
    
- That the  **Recordset** should be restored from a file where it was previously saved.
    
This argument can contain a bitmask of  **CommandTypeEnum** or **ExecuteOptionEnum** values. A **CommandTypeEnum** passed in the _Options_ argument sets the **CommandType** property of the **Recordset**.


 **Note**  The  **ExecuteOpenEnum** values of **adExecuteNoRecords** and **adExecuteStream** cannot be used with **Open**.

If the  **CommandType** property value equals **adCmdUnknown** (the default value), you might experience diminished performance, because ADO must make calls to the provider to determine whether the **CommandText** property is a SQL statement, a stored procedure, or a table name. If you know what type of command you are using, setting the **CommandType** property instructs ADO to go directly to the relevant code. If the **CommandType** property does not match the type of command in the **CommandText** property, an error occurs when you call the **Open** method.

For more information about using these enumerated constants for  _Options_ and with other ADO methods and properties, see CommandTypeEnum and ExecuteOptionEnum.


## ActiveConnection Argument
<a name="sectionSection1"> </a>

You can pass in either a  **Connection** object or a connection string as the _ActiveConnection_ argument.


```
recordset .Open Source, ActiveConnection, CursorType, LockType, Options
```

The  _ActiveConnection_ argument corresponds to the **ActiveConnection** property and specifies in which connection to open the **Recordset** object. If you pass a connection definition for this argument, ADO opens a new connection using the specified parameters. After opening the **Recordset** with a client-side cursor ( **CursorLocation** = **adUseClient** ), you can change the value of this property to send updates to another provider. Or you can set this property to Nothing (in Microsoft Visual Basic) or NULL to disconnect the **Recordset** from any provider. Changing **ActiveConnection** for a server-side cursor generates an error, however.

If you pass a  **Command** object in the _Source_ argument and also pass an _ActiveConnection_ argument, an error occurs because the **ActiveConnection** property of the **Command** object must already be set to a valid **Connection** object or connection string.


## CursorType Argument
<a name="sectionSection2"> </a>


```
recordset .Open Source, ActiveConnection, CursorType, LockType, Options
```

As discussed in The Significance of Cursor Location, the type of cursor that your application uses will determine which capabilities are available to the resultant  **Recordset** (if any). For a detailed examination of cursor types, see Chapter 8: Understanding Cursors and Locks.

The  _CursorType_ argument can accept any of the **CursorTypeEnum** values.


## LockType Argument
<a name="sectionSection3"> </a>


```
recordset .Open Source, ActiveConnection, CursorType, LockType, Options
```

Set the  _LockType_ argument to specify what type of locking the provider should use when opening the **Recordset**. The different types of locking are discussed in Chapter 8: Understanding Cursors and Locks.

The  _LockType_ argument can accept any of the **LockTypeEnum** values.


## Retrieving Multiple Recordsets
<a name="sectionSection4"> </a>

You might occasionally need to execute a command that will return more than one result set. A common example is a stored procedure that runs against a SQL Server database, as in the following example. The stored procedure contains a COMPUTE clause to return the average price of all products in the table. The definition of the stored procedure is as follows:


```sql
 
CREATE PROCEDURE ProductsWithAvgPrice  
AS 
SELECT ProductID, ProductName, UnitPrice  
  FROM PRODUCTS  
  COMPUTE AVG(UnitPrice) 

```

The Microsoft OLE DB Provider for SQL Server returns multiple result sets to ADO when the command contains a COMPUTE clause. Therefore, the ADO code must use the  **NextRecordset** method to access the data in the second result set, as shown here:




```vb
 
'BeginNextRs 
    On Error GoTo ErrHandler: 
     
    Dim objConn As New ADODB.Connection 
    Dim objCmd As New ADODB.Command 
    Dim objRs As New ADODB.Recordset 
 
    Set objConn = GetNewConnection 
    objCmd.ActiveConnection = objConn 
     
    objCmd.CommandText = "ProductsWithAvgPrice" 
    objCmd.CommandType = adCmdStoredProc 
     
    Set objRs = objCmd.Execute 
     
    Do While Not objRs.EOF 
        Debug.Print objRs(0) &; vbTab &; objRs(1) &; vbTab &; _ 
                    objRs(2) 
        objRs.MoveNext 
    Loop 
     
    Set objRs = objRs.NextRecordset 
     
    Debug.Print "AVG. PRICE = $ " &; objRs(0) 
 
    'clean up 
    objRs.Close 
    objConn.Close 
    Set objRs = Nothing 
    Set objConn = Nothing 
    Set objCmd = Nothing 
    Exit Sub 
     
ErrHandler: 
    'clean up 
    If objRs.State = adStateOpen Then 
        objRs.Close 
    End If 
     
    If objConn.State = adStateOpen Then 
        objConn.Close 
    End If 
     
    Set objRs = Nothing 
    Set objConn = Nothing 
    Set objCmd = Nothing 
     
    If Err <> 0 Then 
        MsgBox Err.Source &; "-->" &; Err.Description, , "Error" 
    End If 
'EndNextRs 

```

For more information, see NextRecordset.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

