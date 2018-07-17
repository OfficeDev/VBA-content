---
title: Collections (Visual C++ Syntax Index with import)
ms.prod: access
ms.assetid: 839b8c78-b6dc-ea2b-fe9c-305b8b47b4b9
ms.date: 06/08/2017
---


# Collections (Visual C++ Syntax Index with #import)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Properties Collection](#sectionSection0)
[Errors Collection](#sectionSection1)
[Parameters Collection](#sectionSection2)
[Fields Collection](#sectionSection3)


It is useful to know that collections inherit certain common methods and properties.
All collections inherit the  **Count** property and **Refresh** method, and all collections add the **Item** property. The **Errors** collection adds the **Clear** method. The **Parameters** collection inherits the **Append** and **Delete** methods, while the **Fields** collection adds the **Append**, **Delete**, and **Update** methods.

## Properties Collection
<a name="sectionSection0"> </a>

 **Methods**


```
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 

```

 **Properties**




```
 
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t &; Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 

```


## Errors Collection
<a name="sectionSection1"> </a>

 **Methods**


```
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthclear_HV10294165.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 

```

 **Properties**




```
 
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t &; Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 

```


## Parameters Collection
<a name="sectionSection2"> </a>

 **Methods**


```
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthappend_HV10294078.xml( IDispatch * Object ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcoldelete_HV10294294.xml( const _variant_t &; Index ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 

```

 **Properties**




```
 
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t &; Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 

```


## Fields Collection
<a name="sectionSection3"> </a>

 **Methods**


```c#
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthappend_HV10294078.xml( _bstr_t Name , enum DataTypeEnum Type , long DefinedSize , 
 enum FieldAttributeEnum Attrib , const _variant_t &; FieldValue  = 
 vtMissing ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcoldeletefield_HV10294293.xml( const _variant_t &; Index ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadorefresh_HV10294718.xml( ); 
 
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthupdate_HV10294888.xml( ); 

```

 **Properties**




```
 
long GetCount( ); 
__declspec(property(get=GetCount)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocount_HV10294234.xml; 
 
PropertyPtr GetItem( const _variant_t &; Index ); 
__declspec(property(get=GetItem)) PropertyPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproitem_HV10294463.xml[]; 

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

