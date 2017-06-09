---
title: Recordset (ADO/WFC Syntax)
ms.prod: access
ms.assetid: 28314537-2585-6e29-2014-e7fd8ae78542
ms.date: 06/08/2017
---


# Recordset (ADO/WFC Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **package com.ms.wfc.data**

 **Constructors**



```
 
public Invalid DDUE based on source, error:link not allowed in code, link filename:mdobjodbrec_HV10294709.xml() 
public Recordset(Object r ) 

```

 **Methods**



```js
 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthaddnew_HV10294007.xml(Object[] fieldList , Object[] valueList ) 
public void addNew(Object[] valueList ) 
public void addNew() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcancelbatch_HV10294131.xml(int affectRecords ) 
public void cancelBatch() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcancelupdate_HV10294132.xml() 
public Object Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclone_HV10294166.xml() 
public Object clone(int lockType ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclose_HV10294173.xml() 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcomparebookmarks_HV10294199.xml(Object bookmark1 , Object bookmark2 ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthrstdelete_HV10294295.xml(int affectRecords ) 
public void delete() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthfindmethodado_HV10294381.xml(String criteria ) 
public void find(String criteria , int SkipRecords ) 
public void find(String criteria , int SkipRecords , int searchDirection ) 
public void find(String criteria , int SkipRecords , int searchDirection , Object bmkStart ) 
public Object[][] Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthgetrows_HV10294398.xml(int Rows , Object bmkStart , Object[] fieldList ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmove_HV10294521.xml(int numRecords ) 
public void move(int numRecords,  Object bmkStart ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmovefirst_HV10294526.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmovefirst_HV10294526.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmovefirst_HV10294526.xml() 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmovefirst_HV10294526.xml() 
public Recordset Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthnextrec_HV10294541.xml() 
public Recordset nextRecordset(int[] recordsAffected ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthrstopen_HV10294566.xml() 
public void open(Object source ) 
public void open(Object source , Object activeConnection ) 
public void open(Object source , Object activeConnection , int cursorType ) 
public void open(Object source , Object activeConnection , int cursorType , 
 int lockType )public void open(Object source , Object activeConnection , int cursorType , 
 int lockType , int options ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorequery_HV10294728.xml() 
public void requery(int options ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadoresync_HV10294735.xml() 
public void resync(int affectRecords,  int resyncValues ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthrstsave_HV10294750.xml(String fileName ) 
public void save(String fileName,  int persistFormat ) 
public boolean Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthsupports_HV10294844.xml(int cursorOptions ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthupdate_HV10294888.xml() 
public void update(Object[] valueList ) 
public void update(Object[] fieldList , Object[] valueList ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthupdatebatch_HV10294893.xml() 
public void updateBatch(int affectRecords ) 

```

 **Properties**



```js
 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproabpage_HV10293970.xml() 
public void setAbsolutePage(int page ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproabpos_HV10293979.xml() 
public void setAbsolutePosition(int pos ) 
public Command Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecommand_HV10293982.xml() 
public Connection Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecon_HV10293988.xml() 
public void setActiveConnection(String conn ) 
public void setActiveConnection(com.ms.wfc.data.Connection c ) 
public boolean Invalid DDUE based on source, error:link not allowed in code, link filename:mdprobof_HV10294113.xml() 
public boolean Invalid DDUE based on source, error:link not allowed in code, link filename:mdprobof_HV10294113.xml() 
public Object Invalid DDUE based on source, error:link not allowed in code, link filename:mdprobookmark_HV10294117.xml() 
public void setBookmark(Object bmk ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocachesize_HV10294119.xml() 
public void setCacheSize(int size ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocursorlocation_HV10294254.xml() 
public void setCursorLocation(int cursorLoc ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocursortype_HV10294257.xml() 
public void setCursorType(int cursorType ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodatamember_HV10294276.xml() 
public void setDataMember(String pbstrDataMember ) 
public Iunknown Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodatasource_HV10294277.xml() 
public void setDataSource(IUnknown dataSource ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproeditmode_HV10294326.xml() 
public Object Invalid DDUE based on source, error:link not allowed in code, link filename:mdprofilter_HV10294373.xml() 
public void setFilter(Object filter ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprolocktype_HV10294487.xml() 
public void setLockType(int lockType ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromarshaloptions_HV10294491.xml() 
public void setMarshalOptions(int options ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromaxrecords_HV10294496.xml() 
public void setMaxRecords(int maxRecords ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdpropagecount_HV10294585.xml() 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdpropagesize_HV10294586.xml() 
public void setPageSize(int pageSize ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprorecordcount_HV10294700.xml() 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdprosortpropertyado_HV10294782.xml() 
public void setSort(String criteria ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdprorstsource_HV10294794.xml() 
public void setSource(String query ) 
public void setSource(com.ms.wfc.data.Command command ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml() 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostatus_HV10294810.xml() 
public boolean Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostayinsync_HV10294815.xml() 
public void setStayInSync(boolean pbStayInSync ) 
public com.ms.wfc.data.Field Invalid DDUE based on source, error:link not allowed in code, link filename:mdobjfield_HV10294362.xml(int n ) 
public com.ms.wfc.data.Field getField(String n ) 
public com.ms.wfc.data.Fields Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolfields_HV10294366.xml() 
public AdoProperties Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolproperties_HV10294633.xml() 

```

 **Events**
For more information about ADO/WFC events, see [ADO Event Instantiation by Language](ado-event-instantiation-by-language.md).



```
 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtendofrecordset_HV10294329.xml(RecordsetEventHandler handler ) 
public void removeOnEndOfRecordset(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtfetchcomplete_HV10294356.xml(RecordsetEventHandler handler ) 
public void removeOnFetchComplete(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtfetchprogress_HV10294358.xml(RecordsetEventHandler handler ) 
public void removeOnFetchProgress(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillchangefield_HV10294950.xml(RecordsetEventHandler handler ) 
public void removeOnFieldChangeComplete(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillmove_HV10294955.xml(RecordsetEventHandler handler ) 
public void removeOnMoveComplete(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillchangerecord_HV10294951.xml(RecordsetEventHandler handler ) 
public void removeOnRecordChangeComplete(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillchangerecordset_HV10294952.xml(RecordsetEventHandler handler ) 
public void removeOnRecordsetChangeComplete(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillchangefield_HV10294950.xml(RecordsetEventHandler handler ) 
public void removeOnWillChangeField(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillchangerecord_HV10294951.xml(RecordsetEventHandler handler ) 
public void removeOnWillChangeRecord(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillchangerecordset_HV10294952.xml(RecordsetEventHandler handler ) 
public void removeOnWillChangeRecordset(RecordsetEventHandler handler ) 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdevtwillmove_HV10294955.xml(RecordsetEventHandler handler ) 
public void removeOnWillMove(RecordsetEventHandler handler ) 

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

