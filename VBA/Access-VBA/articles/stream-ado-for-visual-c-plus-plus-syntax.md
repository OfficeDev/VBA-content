---
title: Stream (ADO for Visual C++ Syntax)
ms.prod: access
ms.assetid: e1482f15-9ef6-9485-06c2-1123762afc9f
ms.date: 06/08/2017
---


# Stream (ADO for Visual C++ Syntax)



**Applies to:** Access 2013 | Access 2016

 **Methods**

[Cancel](http://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx)(void)[Close](http://msdn.microsoft.com/library/26a7cced-ebeb-70be-f5de-96a35711bc37%28Office.15%29.aspx)(void)[CopyTo](http://msdn.microsoft.com/library/1c1ab950-51f7-7ecc-ccd8-e689db02f06a%28Office.15%29.aspx)(<em>ADOStream  </em><em>DestStream,<em> LONG _CharNumber</em> = -1)<a href="http://msdn.microsoft.com/library/c167e3b1-c133-ce45-6cee-5a1280a1568f%28Office.15%29.aspx" data-raw-source="[Flush](http://msdn.microsoft.com/library/c167e3b1-c133-ce45-6cee-5a1280a1568f%28Office.15%29.aspx)">Flush</a>(void)<a href="http://msdn.microsoft.com/library/33fd543f-bd24-9199-7540-2889b69221c8%28Office.15%29.aspx" data-raw-source="[LoadFromFile](http://msdn.microsoft.com/library/33fd543f-bd24-9199-7540-2889b69221c8%28Office.15%29.aspx)">LoadFromFile</a>(BSTR <em> FileName</em> )<a href="http://msdn.microsoft.com/library/fa2e6aaa-e9f5-009c-f3a0-050a00abf9b0%28Office.15%29.aspx" data-raw-source="[Open](http://msdn.microsoft.com/library/fa2e6aaa-e9f5-009c-f3a0-050a00abf9b0%28Office.15%29.aspx)">Open</a>(VARIANT <em> Source,</em> ConnectModeEnum <em> Mode,</em> StreamOpenOptionsEnum <em> Options,</em> BSTR <em> UserName,</em> BSTR <em> Password</em> )<a href="http://msdn.microsoft.com/library/91c3ad34-f891-5be0-1fc1-c5c8a2ff07a4%28Office.15%29.aspx" data-raw-source="[Read](http://msdn.microsoft.com/library/91c3ad34-f891-5be0-1fc1-c5c8a2ff07a4%28Office.15%29.aspx)">Read</a>(long <em> NumBytes,</em> VARIANT _</em>pVal_ )[ReadText](http://msdn.microsoft.com/library/08f5bac4-dccd-696c-09a7-e1ba0cb38d79%28Office.15%29.aspx)(long <em> NumChars,</em> BSTR <em>*pbstr</em> )[SaveToFile](http://msdn.microsoft.com/library/db0fd95e-8ef3-af87-5346-8f8713153ca7%28Office.15%29.aspx)(BSTR <em> FileName,</em> SaveOptionsEnum <em>Options</em> =adSaveCreateNotExist)[SetEOS](http://msdn.microsoft.com/library/d438eecf-7ab3-a07d-b6d5-8816db4aae7c%28Office.15%29.aspx)(void)[SkipLine](http://msdn.microsoft.com/library/419c24c3-6b84-eed0-5884-f2dcd485dc3d%28Office.15%29.aspx)(void)[Write](http://msdn.microsoft.com/library/cabe4581-409f-7f05-bd59-d495bfb2c6fd%28Office.15%29.aspx)(VARIANT <em> Buffer</em> )[WriteText](http://msdn.microsoft.com/library/1ca2d9d5-11f4-d088-6fc3-53240208bb09%28Office.15%29.aspx)(BSTR <em> Data,</em> StreamWriteEnum <em>Options</em> =adWriteChar)
 
<strong>Properties</strong>

[get_Charset](http://msdn.microsoft.com/library/454f664e-6d62-eec9-487d-882c2f9503b0%28Office.15%29.aspx)(BSTR  <em>*pbstrCharset</em> ) <strong>put_Charset</strong> (BSTR <em> Charset</em> )[get_EOS](http://msdn.microsoft.com/library/97cd23ef-cca8-4dcc-2641-082a0e1b853c%28Office.15%29.aspx)(VARIANT_BOOL  <em>*pEOS</em> )[get_LineSeparator](http://msdn.microsoft.com/library/9f1323cd-d4ed-2bfa-554b-faebab529548%28Office.15%29.aspx)(LineSeparatorEnum  <em>*pLS</em> ) <strong>put_LineSeparator</strong> (LineSeparatorEnum <em> LineSeparator</em> )[get_Mode](http://msdn.microsoft.com/library/62086f4f-8624-16c4-dae1-a17475d1864d%28Office.15%29.aspx)(ConnectModeEnum  <em>*pMode</em> ) <strong>put_Mode</strong> (ConnectModeEnum <em> Mode</em> )[get_Position](http://msdn.microsoft.com/library/a07c9197-673b-ddf2-fca9-b0b54fbd67b4%28Office.15%29.aspx)(LONG  <em>*pPos</em> ) <strong>put_Position</strong> (LONG <em> Position</em> )[get_Size](size-property-ado-stream.md)(LONG  <em>*pSize</em> )[get_State](http://msdn.microsoft.com/library/ade0a50c-e2d8-23ac-4ea9-b012fedcd5db%28Office.15%29.aspx)(ObjectStateEnum  <em>*pState</em> )[get_Type](http://msdn.microsoft.com/library/43872c74-51bf-47ae-6bdc-55d25b0dc84a%28Office.15%29.aspx)(StreamTypeEnum  <em>*pType</em> ) <strong>put_Type</strong> (StreamTypeEnum <em> Type</em> )
 
<strong>ACCESS SUPPORT RESOURCES</strong><br>

[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>

[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>

[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>

[Search for specific Access error codes on Bing](http://www.bing.com/)<br>

[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>

[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>

[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>

[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

