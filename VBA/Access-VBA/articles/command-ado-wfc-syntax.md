---
title: Command (ADO/WFC Syntax)
ms.prod: access
ms.assetid: fd244794-8831-883a-7892-3ad04d732790
ms.date: 06/08/2017
---


# Command (ADO/WFC Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **package com.ms.wfc.data**

 **Constructor**



```
 
public Invalid DDUE based on source, error:link not allowed in code, link filename:mdobjcommand_HV10294191.xml() 
public Command(String commandtext ) 

```

 **Methods**



```
 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml() 
public com.ms.wfc.data.Parameter Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcreateparam_HV10294243.xml(String 
 Name , int Type , int Direction , int Size , Object Value ) 
public Recordset Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcmdexecute_HV10294344.xml() 
public Recordset execute(Object[] parameters ) 
public Recordset execute(Object[] parameters , int options ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcmdexecute_HV10294344.xml(Object[] parameters ) 
public int executeUpdate(Object[] parameters , int options ) 
public int executeUpdate() 

```

The  **executeUpdate** method is a special case method that calls the underlying ADO **execute** method with certain parameters. The **executeUpdate** method does not support the return of a **Recordset** object, so the **execute** method's _options_ parameter is modified with **AdoEnums.ExecuteOptions.NORECORDS**. After the **execute** method completes, its updated _RecordsAffected_ parameter is passed back to the **executeUpdate** method, which is finally returned as an **int**.
 **Properties**



```js
 
public com.ms.wfc.data.Connection Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecon_HV10293988.xml() 
public void setActiveConnection(com.ms.wfc.data.Connection con ) 
public void setActiveConnection(String conString ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtext_HV10294195.xml() 
public void setCommandText(String command ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtimeout_HV10294196.xml() 
public void setCommandTimeout(int timeout ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtype_HV10294197.xml() 
public void setCommandType(int type ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdproname_HV10294535.xml() 
public void setName(String name ) 
public boolean Invalid DDUE based on source, error:link not allowed in code, link filename:mdproprepared_HV10294617.xml() 
public void setPrepared(boolean prepared ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml() 
public com.ms.wfc.data.Parameter Invalid DDUE based on source, error:link not allowed in code, link filename:mdobjparameter_HV10294590.xml(int n ) 
public com.ms.wfc.data.Parameter getParameter(String n ) 
public com.ms.wfc.data.Parameters Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolparameters_HV10294594.xml() 
public AdoProperties Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolproperties_HV10294633.xml() 

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

