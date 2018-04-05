---
title: ObjectProxy (ADO/WFC Syntax)
ms.prod: access
ms.assetid: 8e3224b7-0b1d-1e08-eaa7-ceb0b6f5411c
ms.date: 06/08/2017
---


# ObjectProxy (ADO/WFC Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Methods](#sectionSection0)
[Syntax](#sectionSection1)
[Returns](#sectionSection2)
[Parameters](#sectionSection3)


An  **ObjectProxy** object represents a server, and is returned by the **createObject** method of the[DataSpace](http://msdn.microsoft.com/library/7db181d5-422b-49fe-b6af-a20f5da520ff%28Office.15%29.aspx) object. The ObjectProxy class has one method, **call**, which can invoke a method on the server and return an object resulting from that invocation.
 **package com.ms.wfc.data**

## Methods
<a name="sectionSection0"> </a>

 **Call Method (ADO/WFC Syntax)**

Invokes a method on the server represented by the ObjectProxy. Optionally, method arguments may be passed as an array of objects.


## Syntax
<a name="sectionSection1"> </a>


```vb
 
public Object ObjectProxy .call( String method  ) 
public Object ObjectProxy .call( String method , Object[] args ) 

```


## Returns
<a name="sectionSection2"> </a>


- Object
    
- An object resulting from invoking the method.
    

## Parameters
<a name="sectionSection3"> </a>


-  _ObjectProxy_
    
- An  **ObjectProxy** object that represents the server.
    
-  _method_
    
- A String, containing the name of the method to invoke on the server.
    
-  _args_
    
- Optional. An array of objects that are arguments to the method on the server. Java data types are automatically converted to data types suitable for use on the server.
    
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

