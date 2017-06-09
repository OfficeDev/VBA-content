---
title: Recordset (Visual C++ Syntax Index with import)
ms.prod: access
ms.assetid: 807e0ce2-2f28-cb4f-41ae-fa4834504a01
ms.date: 06/08/2017
---


# Recordset (Visual C++ Syntax Index with #import)

  

**Applies to:** Access 2013 | Access 2016

 **Methods**




```c#
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthaddnew_HV10294007.xml( const _variant_t &; FieldList  = vtMissing, 
 const _variant_t &; Values  =vtMissing); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcancelbatch_HV10294131.xml( enum AffectEnum AffectRecords ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcancelupdate_HV10294132.xml( ); 
 
_RecordsetPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclone_HV10294166.xml( enum LockTypeEnum LockType ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclose_HV10294173.xml( ); 
 
enum CompareEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcomparebookmarks_HV10294199.xml( const _variant_t 
 &; Bookmark1 , const _variant_t &; Bookmark2  ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthrstdelete_HV10294295.xml( enum AffectEnum AffectRecords ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthfindmethodado_HV10294381.xml( _bstr_t Criteria , long SkipRecords , enum 
 SearchDirectionEnum SearchDirection , const _variant_t &; Start  = 
 vtMissing ); 
 
_variant_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthgetrows_HV10294398.xml( long Rows , const _variant_t &; Start  = 
 vtMissing, const _variant_t &; Fields  = vtMissing ); 
 
_bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthgetstringmethod_recordset_ado_HV10294404.xml( enum 
 StringFormatEnum StringFormat , long NumRows , _bstr_t 
 ColumnDelimeter , _bstr_t RowDelimeter , _bstr_t NullExpr ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmove_HV10294521.xml( long NumRecords , const _variant_t &; Start  = vtMissing); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmovefirst_HV10294526.xml( ); 
HRESULT MoveLast( ); 
HRESULT MoveNext( ); 
HRESULT MovePrevious( ); 
 
_RecordsetPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthnextrec_HV10294541.xml( VARIANT * RecordsAffected ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthrstopen_HV10294566.xml( const _variant_t &; Source , const _variant_t &; 
 ActiveConnection , enum CursorTypeEnum CursorType , enum LockTypeEnum 
 LockType , long Options ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorequery_HV10294728.xml( long Options ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthupdate_HV10294888.xml( const _variant_t &; Fields  = vtMissing, const 
 _variant_t &; Values  =vtMissing); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthupdatebatch_HV10294893.xml( enum AffectEnum AffectRecords ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadoresync_HV10294735.xml( enum AffectEnum AffectRecords , enum 
 ResyncEnum ResyncValues  ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthrstsave_HV10294750.xml( const _variant_t &; Destination , enum 
 PersistFormatEnum PersistFormat  ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthseek_HV10294763.xml( const _variant_t &; KeyValues, enum SeekEnumSeekOption ); 
 
VARIANT_BOOL Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthsupports_HV10294844.xml( enum CursorOptionEnum CursorOptions ); 

```

 **Properties**



```c#
 
enum PositionEnum GetAbsolutePage( ); 
void PutAbsolutePage( enum PositionEnum pl ); 
__declspec(property(get=GetAbsolutePage,put=PutAbsolutePage)) enum 
 PositionEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdproabpage_HV10293970.xml; 
 
enum PositionEnum GetAbsolutePosition( ); 
void PutAbsolutePosition( enum PositionEnum pl ); 
__declspec(property(get=GetAbsolutePosition,put=PutAbsolutePosition)) 
 enum PositionEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdproabpos_HV10293979.xml; 
 
IDispatchPtr GetActiveCommand( ); 
__declspec(property(get=GetActiveCommand)) IDispatchPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecommand_HV10293982.xml; 
 
void Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecon_HV10293988.xml( IDispatch * pvar ); 
void PutActiveConnection( const _variant_t &; pvar ); 
_variant_t GetActiveConnection( ); 
 
enum CursorLocationEnum GetCursorLocation( ); 
void PutCursorLocation( enum CursorLocationEnum plCursorLoc ); 
__declspec(property(get=GetCursorLocation,put=PutCursorLocation)) enum 
 CursorLocationEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocursorlocation_HV10294254.xml; 
 
VARIANT_BOOL GetBOF( ); 
__declspec(property(get=GetBOF)) VARIANT_BOOL Invalid DDUE based on source, error:link not allowed in code, link filename:mdprobof_HV10294113.xml; 
 
VARIANT_BOOL GetEndOfFile( ); // Actually, GetEOF. Renamed in #import. 
__declspec(property(get=GetEndOfFile)) VARIANT_BOOL Invalid DDUE based on source, error:link not allowed in code, link filename:mdprobof_HV10294113.xml; 
 
_variant_t GetBookmark( ); 
void PutBookmark( const _variant_t &; pvBookmark ); 
__declspec(property(get=GetBookmark,put=PutBookmark)) _variant_t 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprobookmark_HV10294117.xml; 
 
long GetCacheSize( ); 
void PutCacheSize( long pl ); 
__declspec(property(get=GetCacheSize,put=PutCacheSize)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocachesize_HV10294119.xml; 
 
enum CursorTypeEnum GetCursorType( ); 
void PutCursorType( enum CursorTypeEnum plCursorType ); 
__declspec(property(get=GetCursorType,put=PutCursorType)) enum 
 CursorTypeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocursortype_HV10294257.xml; 
 
_bstr_t GetDataMember( ); 
void PutDataMember( _bstr_t pbstrDataMember ); 
__declspec(property(get=GetDataMember,put=PutDataMember)) _bstr_t 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodatamember_HV10294276.xml; 
 
IUnknownPtr GetDataSource( ); 
void PutRefDataSource( IUnknown * ppunkDataSource ); 
__declspec(property(get=GetDataSource,put=PutRefDataSource)) IUnknownPtr 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodatasource_HV10294277.xml; 
 
enum EditModeEnum GetEditMode( ); 
__declspec(property(get=GetEditMode)) enum EditModeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdproeditmode_HV10294326.xml; 
 
FieldsPtr GetFields( ); 
__declspec(property(get=GetFields)) FieldsPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolfields_HV10294366.xml; 
 
_variant_t GetFilter( ); 
void PutFilter( const _variant_t &; Criteria ); 
__declspec(property(get=GetFilter,put=PutFilter)) _variant_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdprofilter_HV10294373.xml; 
 
_bstr_t GetIndex( ); 
void PutIndex( _bstr_t pbstrIndex ); 
__declspec(property(get=GetIndex,put=PutIndex)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproindex_HV10294437.xml; 
 
enum LockTypeEnum GetLockType( ); 
void PutLockType( enum LockTypeEnum plLockType ); 
__declspec(property(get=GetLockType,put=PutLockType)) enum LockTypeEnum 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprolocktype_HV10294487.xml; 
 
enum MarshalOptionsEnum GetMarshalOptions( ); 
void PutMarshalOptions( enum MarshalOptionsEnum peMarshal ); 
__declspec(property(get=GetMarshalOptions,put=PutMarshalOptions)) enum 
 MarshalOptionsEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromarshaloptions_HV10294491.xml; 
 
long GetMaxRecords( ); 
void PutMaxRecords( long plMaxRecords ); 
__declspec(property(get=GetMaxRecords,put=PutMaxRecords)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromaxrecords_HV10294496.xml; 
 
long GetPageCount( ); 
__declspec(property(get=GetPageCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdpropagecount_HV10294585.xml; 
 
long GetPageSize( ); 
void PutPageSize( long pl ); 
__declspec(property(get=GetPageSize,put=PutPageSize)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdpropagesize_HV10294586.xml; 
 
long GetRecordCount( ); 
__declspec(property(get=GetRecordCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprorecordcount_HV10294700.xml; 
 
_bstr_t GetSort( ); 
void PutSort( _bstr_t Criteria ); 
__declspec(property(get=GetSort,put=PutSort)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdprosortpropertyado_HV10294782.xml; 
 
void Invalid DDUE based on source, error:link not allowed in code, link filename:mdprorstsource_HV10294794.xml( IDispatch * pvSource ); 
void PutSource( _bstr_t pvSource ); 
_variant_t GetSource( ); 
 
long GetState( ); 
__declspec(property(get=GetState)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml; 
 
long GetStatus( ); 
__declspec(property(get=GetStatus)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostatus_HV10294810.xml; 
 
VARIANT_BOOL GetStayInSync( ); 
void PutStayInSync( VARIANT_BOOL pbStayInSync ); 
__declspec(property(get=GetStayInSync,put=PutStayInSync)) VARIANT_BOOL 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostayinsync_HV10294815.xml; 

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

