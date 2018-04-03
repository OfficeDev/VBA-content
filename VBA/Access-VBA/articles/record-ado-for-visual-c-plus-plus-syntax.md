---
title: Record (ADO for Visual C++ Syntax)
ms.prod: access
ms.assetid: e9a1300e-e2d8-7ad9-e0d6-61be720b83af
ms.date: 06/08/2017
---


# Record (ADO for Visual C++ Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **Methods**

[Cancel](http://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx)(void)[Close](http://msdn.microsoft.com/library/26a7cced-ebeb-70be-f5de-96a35711bc37%28Office.15%29.aspx)(void)[CopyRecord](http://msdn.microsoft.com/library/724e4358-f216-8e47-5bab-c72770ece5a4%28Office.15%29.aspx)(BSTR _ Source,_ BSTR _ Destination,_ BSTR _ UserName,_ BSTR _ Password,_ CopyRecordOptionsEnum _ Options,_ VARIANT_BOOL _ Async,_ BSTR _*pbstrNewURL_ )[DeleteRecord](http://msdn.microsoft.com/library/ba71187f-e580-bba8-f41b-bedfa0bc2b04%28Office.15%29.aspx)(BSTR _ Source,_ VARIANT_BOOL _ Async_ )[GetChildren](http://msdn.microsoft.com/library/998cf640-ffc7-51e1-4d1e-4797f7cdea4a%28Office.15%29.aspx)(_ADORecordset * _*ppRSet_ )[MoveRecord](http://msdn.microsoft.com/library/efc341a2-0e08-a838-5925-8d4c46377e48%28Office.15%29.aspx)(BSTR _ Source,_ BSTR _ Destination,_ BSTR _ UserName,_ BSTR _ Password,_ MoveRecordOptionsEnum _ Options,_ VARIANT_BOOL _ Async,_ BSTR _*pbstrNewURL_ )[Open](http://msdn.microsoft.com/library/ba71c5c7-326e-d3b6-0e74-e8343ee6896f%28Office.15%29.aspx)(VARIANT _ Source,_ VARIANT _ ActiveConnection,_ ConnectModeEnum _ Mode,_ RecordCreateOptionsEnum _ CreateOptions,_ RecordOpenOptionsEnum _Options,_ BSTR _ UserName,_ BSTR _ Password_ )
 **Properties**
[get_ActiveConnection](http://msdn.microsoft.com/library/5501b2d7-b62c-5fff-1edd-2b7efb3f8c4a%28Office.15%29.aspx)(VARIANT  _*pvar_ ) **put_ActiveConnection** (BSTR _ bstrConn_ ) **putref_ActiveConnection** (_ADOConnection _*Con_ )[get_Fields](http://msdn.microsoft.com/library/029aa738-8726-54a6-1813-b152813948bc%28Office.15%29.aspx)(ADOFields * _*ppFlds_ )[get_Mode](http://msdn.microsoft.com/library/62086f4f-8624-16c4-dae1-a17475d1864d%28Office.15%29.aspx)(ConnectModeEnum  _*pMode_ ) **put_Mode** (ConnectModeEnum _ Mode_ )[get_ParentURL](http://msdn.microsoft.com/library/ec7ec476-6f9e-8486-fe02-74995975df5c%28Office.15%29.aspx)(BSTR  _*pbstrParentURL_ )[get_RecordType](http://msdn.microsoft.com/library/a42001a6-7312-162d-dd71-c82f8c9d527f%28Office.15%29.aspx)(RecordTypeEnum  _*pType_ )[get_Source](http://msdn.microsoft.com/library/f36f0f5f-4493-d8c5-db4b-c72f5031bcb3%28Office.15%29.aspx)(VARIANT  _*pvar_ ) **put_Source** (BSTR _ Source_ ) **putref_Source** (IDispatch _*Source_ )[get_State](http://msdn.microsoft.com/library/ade0a50c-e2d8-23ac-4ea9-b012fedcd5db%28Office.15%29.aspx)(ObjectStateEnum  _*pState_ )
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

