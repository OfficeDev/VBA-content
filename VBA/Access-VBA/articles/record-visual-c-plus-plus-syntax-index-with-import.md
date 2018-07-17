---
title: Record (Visual C++ Syntax Index with import)
ms.prod: access
ms.assetid: 87c6d242-4977-2e81-c829-227e6dd326e5
ms.date: 06/08/2017
---


# Record (Visual C++ Syntax Index with #import)

  

**Applies to:** Access 2013 | Access 2016

 **Methods**




```c#
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclose_HV10294173.xml( ); 
 
_bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcopyrecord_HV10294230.xml( _bstr_t Source , _bstr_t Destination , 
 _bstr_t UserName , _bstr_t Password , enum CopyRecordOptionsEnum 
 Options , VARIANT_BOOL Async ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthdeleterecord_HV10294302.xml( _bstr_t Source , VARIANT_BOOL Async ); 
 
_RecordsetPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthgetchildrenmethodado_HV10294390.xml( ); 
 
_bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthmoverecord_HV10294532.xml( _bstr_t Source , _bstr_t Destination , 
 _bstr_t UserName , _bstr_t Password , enum MoveRecordOptionsEnum 
 Options , VARIANT_BOOL Async ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthopenrecord_HV10294565.xml( const _variant_t &; Source , const _variant_t 
&; ActiveConnection , enum ConnectModeEnum Mode , enum 
 RecordCreateOptionsEnum CreateOptions , enum RecordOpenOptionsEnum 
 Options , _bstr_t UserName , _bstr_t Password ); 

```

 **Properties**



```c#
 
_variant_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecon_HV10293988.xml( ); 
void PutActiveConnection( _bstr_t pvar ); 
void PutRefActiveConnection( struct _Connection * pvar ); 
 
FieldsPtr GetFields( ); 
__declspec(property(get=GetFields)) FieldsPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolfields_HV10294366.xml; 
 
enum ConnectModeEnum GetMode( ); 
void PutMode( enum ConnectModeEnum pMode ); 
__declspec(property(get=GetMode,put=PutMode)) enum ConnectModeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromode_HV10294518.xml; 
 
_bstr_t GetParentURL( ); 
__declspec(property(get=GetParentURL)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproparenturl_HV10294604.xml; 
 
enum RecordTypeEnum GetRecordType( ); 
__declspec(property(get=GetRecordType)) enum RecordTypeEnum 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprorecordtypeproperty_HV10294716.xml; 
 
_variant_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdprosourcerecord_HV10294793.xml( ); 
void PutSource( _bstr_t pvar ); 
void PutRefSource( IDispatch * pvar ); 
 
enum ObjectStateEnum GetState( ); 
__declspec(property(get=GetState)) enum ObjectStateEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml; 

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

