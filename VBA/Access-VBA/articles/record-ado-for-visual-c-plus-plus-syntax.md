---
title: Record (ADO for Visual C++ Syntax)
ms.prod: access
ms.assetid: e9a1300e-e2d8-7ad9-e0d6-61be720b83af
ms.date: 06/08/2017
---


# Record (ADO for Visual C++ Syntax)



**Applies to:** Access 2013 | Access 2016

 **Methods**

[Cancel](http://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx)(void)[Close](http://msdn.microsoft.com/library/26a7cced-ebeb-70be-f5de-96a35711bc37%28Office.15%29.aspx)(void)[CopyRecord](http://msdn.microsoft.com/library/724e4358-f216-8e47-5bab-c72770ece5a4%28Office.15%29.aspx)(BSTR <em> Source,</em> BSTR <em> Destination,</em> BSTR <em> UserName,</em> BSTR <em> Password,</em> CopyRecordOptionsEnum <em> Options,</em> VARIANT_BOOL <em> Async,</em> BSTR <em>*pbstrNewURL</em> )[DeleteRecord](http://msdn.microsoft.com/library/ba71187f-e580-bba8-f41b-bedfa0bc2b04%28Office.15%29.aspx)(BSTR <em> Source,</em> VARIANT_BOOL <em> Async</em> )[GetChildren](http://msdn.microsoft.com/library/998cf640-ffc7-51e1-4d1e-4797f7cdea4a%28Office.15%29.aspx)(<em>ADORecordset * </em><em>ppRSet_ )<a href="http://msdn.microsoft.com/library/efc341a2-0e08-a838-5925-8d4c46377e48%28Office.15%29.aspx" data-raw-source="[MoveRecord](http://msdn.microsoft.com/library/efc341a2-0e08-a838-5925-8d4c46377e48%28Office.15%29.aspx)">MoveRecord</a>(BSTR <em> Source,</em> BSTR <em> Destination,</em> BSTR <em> UserName,</em> BSTR <em> Password,</em> MoveRecordOptionsEnum <em> Options,</em> VARIANT_BOOL <em> Async,</em> BSTR _</em>pbstrNewURL_ )[Open](http://msdn.microsoft.com/library/ba71c5c7-326e-d3b6-0e74-e8343ee6896f%28Office.15%29.aspx)(VARIANT <em> Source,</em> VARIANT <em> ActiveConnection,</em> ConnectModeEnum <em> Mode,</em> RecordCreateOptionsEnum <em> CreateOptions,</em> RecordOpenOptionsEnum <em>Options,</em> BSTR <em> UserName,</em> BSTR <em> Password</em> )
 
<strong>Properties</strong>

[get_ActiveConnection](http://msdn.microsoft.com/library/5501b2d7-b62c-5fff-1edd-2b7efb3f8c4a%28Office.15%29.aspx)(VARIANT  <em>*pvar</em> ) <strong>put_ActiveConnection</strong> (BSTR <em> bstrConn</em> ) <strong>putref_ActiveConnection</strong> (<em>ADOConnection </em><em>Con_ )<a href="http://msdn.microsoft.com/library/029aa738-8726-54a6-1813-b152813948bc%28Office.15%29.aspx" data-raw-source="[get_Fields](http://msdn.microsoft.com/library/029aa738-8726-54a6-1813-b152813948bc%28Office.15%29.aspx)">get_Fields</a>(ADOFields * _</em>ppFlds_ )[get_Mode](http://msdn.microsoft.com/library/62086f4f-8624-16c4-dae1-a17475d1864d%28Office.15%29.aspx)(ConnectModeEnum  <em>*pMode</em> ) <strong>put_Mode</strong> (ConnectModeEnum <em> Mode</em> )[get_ParentURL](http://msdn.microsoft.com/library/ec7ec476-6f9e-8486-fe02-74995975df5c%28Office.15%29.aspx)(BSTR  <em>*pbstrParentURL</em> )[get_RecordType](http://msdn.microsoft.com/library/a42001a6-7312-162d-dd71-c82f8c9d527f%28Office.15%29.aspx)(RecordTypeEnum  <em>*pType</em> )[get_Source](http://msdn.microsoft.com/library/f36f0f5f-4493-d8c5-db4b-c72f5031bcb3%28Office.15%29.aspx)(VARIANT  <em>*pvar</em> ) <strong>put_Source</strong> (BSTR <em> Source</em> ) <strong>putref_Source</strong> (IDispatch <em>*Source</em> )[get_State](http://msdn.microsoft.com/library/ade0a50c-e2d8-23ac-4ea9-b012fedcd5db%28Office.15%29.aspx)(ObjectStateEnum  <em>*pState</em> )
 
<strong>ACCESS SUPPORT RESOURCES</strong><br>

[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>

[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>

[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>

[Search for specific Access error codes on Bing](http://www.bing.com/)<br>

[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>

[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>

[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>

[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

