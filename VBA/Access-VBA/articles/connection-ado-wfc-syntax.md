---
title: Connection (ADO/WFC Syntax)
ms.prod: access
ms.assetid: adead04c-7a49-40b8-6d15-5d19c559b1b2
ms.date: 06/08/2017
---


# Connection (ADO/WFC Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **package com.ms.wfc.data**

 **Constructor**



```
 
public Invalid DDUE based on source, error:link not allowed in code, link filename:mdobjconnection_HV10294216.xml() 
public Connection(String connectionstring ) 

```

 **Methods**



```
 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthbegintrans_HV10294108.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthbegintrans_HV10294108.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthbegintrans_HV10294108.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclose_HV10294173.xml() 
public com.ms.wfc.data.Recordset Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcnnexecute_HV10294345.xml(String commandText ) 
public com.ms.wfc.data.Recordset execute(String commandText , int options ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcnnexecute_HV10294345.xml(String commandText ) 
public int executeUpdate(String commandText , int options ) 

```

The  **executeUpdate** method is a special case method that calls the underlying ADO **execute** method with certain parameters. The **executeUpdate** method does not support the return of a **Recordset** object, so the **execute** method's _options_ parameter is modified with **AdoEnums.ExecuteOptions.NORECORDS**. After the **execute** method completes, its updated _RecordsAffected_ parameter is passed back to the **executeUpdate** method, which is finally returned as an **int**.



```
 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcnnopen_HV10294563.xml() 
public void open(String connectionString ) 
public void open(String connectionString , String userID ) 
public void open(String connectionString , String userID , String password ) 
public void open(String connectionString , String userID , String password , int options ) 
public Recordset Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthopenschema_HV10294568.xml(int schema, Object[] 
 restrictions , String schemaID ) 
public Recordset openSchema(int schema) 
public Recordset openSchema(int schema, Object[] restrictions ) 

```

 **Properties**



```
 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproattributes_HV10294098.xml() 
public void setAttributes(int attr ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtimeout_HV10294196.xml() 
public void setCommandTimeout(int timeout ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdproconnectionstring_HV10294218.xml() 
public void setConnectionString(String con ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproconnectiontimeout_HV10294222.xml() 
public void setConnectionTimeout(int timeout ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocursorlocation_HV10294254.xml() 
public void setCursorLocation(int cursorLoc ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodefaultdatabase_HV10294288.xml() 
public void setDefaultDatabase(String db ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproisolationlevel_HV10294459.xml() 
public void setIsolationLevel(int level ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromode_HV10294518.xml() 
public void setMode(int mode ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdproprovider_HV10294673.xml() 
public void setProvider(String provider ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml() 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdproversion_HV10294926.xml() 
public AdoProperties Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolproperties_HV10294633.xml() 
public com.ms.wfc.data.Errors Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolerrors_HV10294338.xml() 

```

 **Events**
For more information about ADO/WFC events, see [ADO Event Instantiation by Language](ado-event-instantiation-by-language.md).



```
 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtbegintranscomplete_HV10294112.xml(ConnectionEventHandler handler ) 
public void removeOnBeginTransComplete(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtbegintranscomplete_HV10294112.xml(ConnectionEventHandler handler ) 
public void removeOnCommitTransComplete(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtconnectconnectionevents_HV10294210.xml(ConnectionEventHandler handler ) 
public void removeOnConnectComplete(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtconnectconnectionevents_HV10294210.xml(ConnectionEventHandler handler ) 
public void removeOnDisconnect(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtexecutecomplete_HV10294351.xml(ConnectionEventHandler handler ) 
public void removeOnExecuteComplete(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtinfomessage_HV10294445.xml(ConnectionEventHandler handler ) 
public void removeOnInfoMessage(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtbegintranscomplete_HV10294112.xml(ConnectionEventHandler handler ) 
public void removeOnRollbackTransComplete(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillconnectevent_HV10294953.xml(ConnectionEventHandler handler ) 
public void removeOnWillConnect(ConnectionEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillexecuteevent_HV10294954.xml(ConnectionEventHandler handler ) 
public void removeOnWillExecute(ConnectionEventHandler handler ) 

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

