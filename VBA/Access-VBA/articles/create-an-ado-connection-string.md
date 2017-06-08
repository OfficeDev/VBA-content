---
title: Create an ADO Connection String
ms.prod: access
ms.assetid: ac29e820-ffbf-a15b-e13d-c9190dfad6ab
ms.date: 06/08/2017
---


# Create an ADO Connection String

To connect to a data source, you must specify a connection string, the parameters of which might differ for each provider and data source. ADO directly supports five arguments in a connection string. Other arguments are passed to the provider that is named in the  _Provider_ argument without any processing by ADO.



|**Argument**|**Description**|
|:-----|:-----|
| _Provider_|Specifies the name of a provider to use for the connection.|
| _File Name_|Specifies the name of a provider-specific file (for example, a persisted data source object) containing preset connection information.|
| _URL_|Specifies the connection string as an absolute URL identifying a resource, such as a file or directory.|
| _Remote Provider_|Specifies the name of a provider to use when opening a client-side connection. (Remote Data Service only.)|
| _Remote Server_|Specifies the path name of the server to use when opening a client-side connection. (Remote Data Service only.)|

The following example 




```
m_sConnStr = "Provider='SQLOLEDB';Data Source='MySqlServer';" &; _ 
 "Initial Catalog='Northwind';Integrated Security='SSPI';"
```

The only ADO parameter supplied in this connection string was "Provider=SQLOLEDB", which indicated the Microsoft OLE DB Provider for SQL Server. Other valid parameters that can be passed in the connection string can be determined by referring to individual providers' documentation.
To open the connection, simply pass the connection string as the first argument in the  **Connection** object's **Open** method:



```
objConn.Open m_sConnStr
```

It is also possible to supply much of this information by setting properties of the  **Connection** object before opening the connection. For example, you could achieve the same effect as the connection string above by using the following code:



```vb
With objConn 
 .Provider = "SQLOLEDB" 
 .DefaultDatabase = "Northwind" 
 .Properties("Data Source") = "MySqlServer" 
 .Properties("Integrated Security") = "SSPI" 
 .Open 
End With 

```


