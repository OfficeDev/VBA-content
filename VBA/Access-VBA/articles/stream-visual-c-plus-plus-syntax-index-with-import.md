---
title: Stream (Visual C++ Syntax Index with import)
ms.prod: access
ms.assetid: a3188858-9c0d-aff6-c893-2111aee77383
ms.date: 06/08/2017
---


# Stream (Visual C++ Syntax Index with #import)

  

**Applies to:** Access 2013 | Access 2016

 **Methods**




```c#
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclose_HV10294173.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcopyto_HV10294233.xml( struct _Stream * DestStream , int CharNumber ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthflush_HV10294386.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthloadfromfile_HV10294485.xml( _bstr_t FileName ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthopenstream_HV10294567.xml( const _variant_t &; Source , enum 
 ConnectModeEnum Mode , enum StreamOpenOptionsEnum Options , _bstr_t 
 UserName , _bstr_t Password ); 
 
_variant_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthread_HV10294692.xml( long NumBytes ); 
 
_bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthreadtext_HV10294694.xml( long NumChars ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthsavetofile_HV10294752.xml( _bstr_t FileName , enum SaveOptionsEnumOptions ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthseteos_HV10294771.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthskipline_HV10294780.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthwrite_HV10294958.xml( const _variant_t &; Buffer ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthwritetext_HV10294959.xml( _bstr_t Data , enum StreamWriteEnumOptions ); 

```

 **Properties**



```c#
 
_bstr_t GetCharset( ); 
void PutCharset( _bstr_t pbstrCharset ); 
__declspec(property(get=GetCharset,put=PutCharset)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocharset_HV10294162.xml; 
 
VARIANT_BOOL GetEOS( ); 
__declspec(property(get=GetEOS)) VARIANT_BOOL Invalid DDUE based on source, error:link not allowed in code, link filename:mdproeos_HV10294332.xml; 
 
enum LineSeparatorEnum GetLineSeparator( ); 
void PutLineSeparator( enum LineSeparatorEnum pLS ); 
__declspec(property(get=GetLineSeparator,put=PutLineSeparator)) enum 
 LineSeparatorEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprolineseparator_HV10294483.xml; 
 
enum ConnectModeEnum GetMode( ); 
void PutMode( enum ConnectModeEnum pMode ); 
__declspec(property(get=GetMode,put=PutMode)) enum ConnectModeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdpromode_HV10294518.xml; 
 
long GetPosition( ); 
void PutPosition( long pPos ); 
__declspec(property(get=GetPosition,put=PutPosition)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdproposition_HV10294611.xml; 
 
long GetSize( ); 
__declspec(property(get=GetSize)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprosizestream_HV10294778.xml; 
 
enum ObjectStateEnum GetState( ); 
__declspec(property(get=GetState)) enum ObjectStateEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml; 
 
enum StreamTypeEnum GetType( ); 
void PutType( enum StreamTypeEnum ptype ); 
__declspec(property(get=GetType,put=PutType)) enum StreamTypeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprotypestream_HV10294865.xml; 

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

