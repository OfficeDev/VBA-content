---
title: Command (ADO for Visual C++ Syntax)
ms.prod: access
ms.assetid: a397daf5-2bcd-6c1a-3fb6-667c1309d0e3
ms.date: 06/08/2017
---


# Command (ADO for Visual C++ Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **Methods**

[Cancel](http://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx)(void)[CreateParameter](http://msdn.microsoft.com/library/cf080a0b-75d2-dcdf-2715-10af147358e9%28Office.15%29.aspx)(BSTR  _Name,_ DataTypeEnum _Type,_ ParameterDirectionEnum _Direction,_ long _Size,_ VARIANT _Value,_ _ADOParameter ** _ppiprm_ )[Execute](execute-method-ado-command.md)(VARIANT * _RecordsAffected,_ VARIANT * _Parameters,_ long _Options,_ _ADORecordset ** _ppirs_ )
 **Properties**
[get_ActiveConnection](http://msdn.microsoft.com/library/5501b2d7-b62c-5fff-1edd-2b7efb3f8c4a%28Office.15%29.aspx)(_ADOConnection ** _ppvObject_ ) **put_ActiveConnection** (VARIANT _vConn_ ) **putref_ActiveConnection** (_ADOConnection * _pCon_ )[get_CommandText](http://msdn.microsoft.com/library/0debec1c-068f-0aea-fce8-e61aa39c5907%28Office.15%29.aspx)(BSTR * _pbstr_ ) **put_CommandText** (BSTR _bstr_ )[get_CommandTimeout](http://msdn.microsoft.com/library/a0b6209c-9feb-08ae-002a-15d1d20734a8%28Office.15%29.aspx)(LONG * _pl_ ) **put_CommandTimeout** (LONG _Timeout_ )[get_CommandType](http://msdn.microsoft.com/library/c8d4fc1c-502b-11f3-af9d-605a03b6f056%28Office.15%29.aspx)(CommandTypeEnum * _plCmdType_ ) **put_CommandType** (CommandTypeEnum _lCmdType_ )[get_Name](http://msdn.microsoft.com/library/4b19bd08-ac3c-86f0-471d-06a37a0d4f89%28Office.15%29.aspx)(BSTR * _pbstrName_ ) **put_Name** (BSTR _bstrName_ )[get_Prepared](http://msdn.microsoft.com/library/33becda2-faab-5000-8904-6ffd8c5805f2%28Office.15%29.aspx)(VARIANT_BOOL * _pfPrepared_ ) **put_Prepared** (VARIANT_BOOL _fPrepared_ )[get_State](http://msdn.microsoft.com/library/ade0a50c-e2d8-23ac-4ea9-b012fedcd5db%28Office.15%29.aspx)(LONG * _plObjState_ )[get_Parameters](http://msdn.microsoft.com/library/554387c3-3572-5391-3b24-c7d3443844cd%28Office.15%29.aspx)(ADOParameters ** _ppvObject_ )
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

