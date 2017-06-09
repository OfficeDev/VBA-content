---
title: Connection (ADO for Visual C++ Syntax)
ms.prod: access
ms.assetid: 04ec8840-a841-1e94-e606-f1c1fb190533
ms.date: 06/08/2017
---


# Connection (ADO for Visual C++ Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **Methods**

[BeginTrans](http://msdn.microsoft.com/library/9a0415f0-9424-8d1c-4779-92e932292d46%28Office.15%29.aspx)(long * _TransactionLevel_ )[CommitTrans](http://msdn.microsoft.com/library/9a0415f0-9424-8d1c-4779-92e932292d46%28Office.15%29.aspx)(void)[RollbackTrans](http://msdn.microsoft.com/library/9a0415f0-9424-8d1c-4779-92e932292d46%28Office.15%29.aspx)(void)[Cancel](http://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx)(void)[Close](http://msdn.microsoft.com/library/26a7cced-ebeb-70be-f5de-96a35711bc37%28Office.15%29.aspx)(void)[Execute](execute-method-ado-connection.md)(BSTR  _CommandText,_ VARIANT * _RecordsAffected,_ long _Options,_ _ADORecordset ** _ppiRset_ )[Open](http://msdn.microsoft.com/library/1adaa17d-dfe1-22e0-3415-720516d138f8%28Office.15%29.aspx)(BSTR  _ConnectionString,_ BSTR _UserID,_ BSTR _Password,_ long _Options_ )[OpenSchema](http://msdn.microsoft.com/library/57771163-a14e-207a-2942-849acb79a9a1%28Office.15%29.aspx)(SchemaEnum  _Schema,_ VARIANT _Restrictions,_ VARIANT _SchemaID,_ _ADORecordset ** _pprset_ )
 **Properties**
[get_Attributes](http://msdn.microsoft.com/library/4cc1f036-606e-7d4b-d270-af374e9d99fa%28Office.15%29.aspx)(long * _plAttr_ ) **put_Attributes** (long _lAttr_ )[get_CommandTimeout](http://msdn.microsoft.com/library/a0b6209c-9feb-08ae-002a-15d1d20734a8%28Office.15%29.aspx)(LONG * _plTimeout_ ) **put_CommandTimeout** (LONG _lTimeout_ )[get_ConnectionString](http://msdn.microsoft.com/library/c67a7daf-258f-d99d-6475-a4aa98d1e99d%28Office.15%29.aspx)(BSTR * _pbstr_ ) **put_ConnectionString** (BSTR _bstr_ )[get_ConnectionTimeout](http://msdn.microsoft.com/library/efc39fd8-afce-5ac0-2fff-cbb55c1a444d%28Office.15%29.aspx)(LONG * _plTimeout_ ) **put_ConnectionTimeout** (LONG _lTimeout_ )[get_CursorLocation](http://msdn.microsoft.com/library/8a048bd4-ae25-a555-1c07-14364b7e6560%28Office.15%29.aspx)(CursorLocationEnum * _plCursorLoc_ ) **put_CursorLocation** (CursorLocationEnum _lCursorLoc_ )[get_DefaultDatabase](http://msdn.microsoft.com/library/a35c5631-f9d9-e51f-950b-e52169830d94%28Office.15%29.aspx)(BSTR * _pbstr_ ) **put_DefaultDatabase** (BSTR _bstr_ )[get_IsolationLevel](http://msdn.microsoft.com/library/19461be5-c94b-4b61-ce08-7abdf702c3dc%28Office.15%29.aspx)(IsolationLevelEnum * _Level_ ) **put_IsolationLevel** (IsolationLevelEnum _Level_ )[get_Mode](http://msdn.microsoft.com/library/62086f4f-8624-16c4-dae1-a17475d1864d%28Office.15%29.aspx)(ConnectModeEnum * _plMode_ ) **put_Mode** (ConnectModeEnum _lMode_ )[get_Provider](http://msdn.microsoft.com/library/1b795f51-93d7-431c-b1fe-0db95f69a56a%28Office.15%29.aspx)(BSTR * _pbstr_ ) **put_Provider** (BSTR _Provider_ )[get_State](http://msdn.microsoft.com/library/ade0a50c-e2d8-23ac-4ea9-b012fedcd5db%28Office.15%29.aspx)(LONG * _plObjState_ )[get_Version](http://msdn.microsoft.com/library/61466895-0a6c-533c-bd93-0ab6af654f24%28Office.15%29.aspx)(BSTR * _pbstr_ )[get_Errors](http://msdn.microsoft.com/library/76c234b8-7fec-11c5-275e-864d5d880ee7%28Office.15%29.aspx)(ADOErrors ** _ppvObject_ )
 **Events**
[BeginTransComplete](http://msdn.microsoft.com/library/9d0ae38e-530a-7a89-a344-f3ab401c2e35%28Office.15%29.aspx)(LONG  _TransactionLevel,_ ADOError * _pError,_ EventStatusEnum * _adStatus,_ _ADOConnection * _pConnection_ )[CommitTransComplete](http://msdn.microsoft.com/library/9d0ae38e-530a-7a89-a344-f3ab401c2e35%28Office.15%29.aspx)(ADOError * _pError,_ EventStatusEnum * _adStatus,_ _ADOConnection * _pConnection_ )[ConnectComplete](http://msdn.microsoft.com/library/8ecb080b-7fc9-7565-25bd-bd57b983750d%28Office.15%29.aspx)(ADOError * _pError,_ EventStatusEnum * _adStatus,_ _ADOConnection * _pConnection_ )[Disconnect](http://msdn.microsoft.com/library/8ecb080b-7fc9-7565-25bd-bd57b983750d%28Office.15%29.aspx)(EventStatusEnum * _adStatus,_ _ADOConnection * _pConnection_ )[ExecuteComplete](http://msdn.microsoft.com/library/47317d97-e373-32f4-9438-2dff46b8d367%28Office.15%29.aspx)(LONG  _RecordsAffected,_ ADOError * _pError,_ EventStatusEnum * _adStatus,_ _ADOCommand * _pCommand,_ _ADORecordset * _pRecordset,_ _ADOConnection * _pConnection_ )[InfoMessage](http://msdn.microsoft.com/library/5d4f487f-96c8-4cf6-60ab-583510d3096f%28Office.15%29.aspx)(ADOError * _pError,_ EventStatusEnum * _adStatus,_ _ADOConnection * _pConnection_ )[RollbackTransComplete](http://msdn.microsoft.com/library/9d0ae38e-530a-7a89-a344-f3ab401c2e35%28Office.15%29.aspx)(ADOError * _pError,_ EventStatusEnum * _adStatus,_ _ADOConnection * _pConnection_ )[WillConnect](http://msdn.microsoft.com/library/8b0e9955-4e7a-7af8-ce6c-7a4ba569a5bb%28Office.15%29.aspx)(BSTR * _ConnectionString,_ BSTR * _UserID,_ BSTR * _Password,_ long * _Options,_ EventStatusEnum * _adStatus,_ _ADOConnection * _pConnection_ )[WillExecute](http://msdn.microsoft.com/library/9f516bfd-246d-9817-4ca3-64598ab466f7%28Office.15%29.aspx)(BSTR * _Source,_ CursorTypeEnum * _CursorType,_ LockTypeEnum * _LockType,_ long * _Options,_ EventStatusEnum * _adStatus,_ _ADOCommand * _pCommand,_ _ADORecordset * _pRecordset,_ _ADOConnection * _pConnection_ )
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

