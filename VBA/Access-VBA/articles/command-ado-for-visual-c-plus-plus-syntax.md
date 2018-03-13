---
title: Command (ADO for Visual C++ Syntax)
ms.prod: access
ms.assetid: a397daf5-2bcd-6c1a-3fb6-667c1309d0e3
ms.date: 06/08/2017
---


# Command (ADO for Visual C++ Syntax)



**Applies to:** Access 2013 | Access 2016

 **Methods**

[Cancel](http://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx)(void)[CreateParameter](http://msdn.microsoft.com/library/cf080a0b-75d2-dcdf-2715-10af147358e9%28Office.15%29.aspx)(BSTR  <em>Name,</em> DataTypeEnum <em>Type,</em> ParameterDirectionEnum <em>Direction,</em> long <em>Size,</em> VARIANT <em>Value,</em> <em>ADOParameter ** _ppiprm</em> )[Execute](execute-method-ado-command.md)(VARIANT * <em>RecordsAffected,</em> VARIANT * <em>Parameters,</em> long <em>Options,</em> <em>ADORecordset ** _ppirs</em> )
 
<strong>Properties</strong>

[get_ActiveConnection](http://msdn.microsoft.com/library/5501b2d7-b62c-5fff-1edd-2b7efb3f8c4a%28Office.15%29.aspx)(<em>ADOConnection ** _ppvObject</em> ) <strong>put_ActiveConnection</strong> (VARIANT <em>vConn</em> ) <strong>putref_ActiveConnection</strong> (<em>ADOConnection * _pCon</em> )[get_CommandText](http://msdn.microsoft.com/library/0debec1c-068f-0aea-fce8-e61aa39c5907%28Office.15%29.aspx)(BSTR * <em>pbstr</em> ) <strong>put_CommandText</strong> (BSTR <em>bstr</em> )[get_CommandTimeout](http://msdn.microsoft.com/library/a0b6209c-9feb-08ae-002a-15d1d20734a8%28Office.15%29.aspx)(LONG * <em>pl</em> ) <strong>put_CommandTimeout</strong> (LONG <em>Timeout</em> )[get_CommandType](http://msdn.microsoft.com/library/c8d4fc1c-502b-11f3-af9d-605a03b6f056%28Office.15%29.aspx)(CommandTypeEnum * <em>plCmdType</em> ) <strong>put_CommandType</strong> (CommandTypeEnum <em>lCmdType</em> )[get_Name](http://msdn.microsoft.com/library/4b19bd08-ac3c-86f0-471d-06a37a0d4f89%28Office.15%29.aspx)(BSTR * <em>pbstrName</em> ) <strong>put_Name</strong> (BSTR <em>bstrName</em> )[get_Prepared](http://msdn.microsoft.com/library/33becda2-faab-5000-8904-6ffd8c5805f2%28Office.15%29.aspx)(VARIANT_BOOL * <em>pfPrepared</em> ) <strong>put_Prepared</strong> (VARIANT_BOOL <em>fPrepared</em> )[get_State](http://msdn.microsoft.com/library/ade0a50c-e2d8-23ac-4ea9-b012fedcd5db%28Office.15%29.aspx)(LONG * <em>plObjState</em> )[get_Parameters](http://msdn.microsoft.com/library/554387c3-3572-5391-3b24-c7d3443844cd%28Office.15%29.aspx)(ADOParameters ** <em>ppvObject</em> )
 
<strong>ACCESS SUPPORT RESOURCES</strong><br>

[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>

[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>

[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>

[Search for specific Access error codes on Bing](http://www.bing.com/)<br>

[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>

[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>

[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>

[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

