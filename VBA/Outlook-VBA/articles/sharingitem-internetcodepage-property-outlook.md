---
title: SharingItem.InternetCodepage Property (Outlook)
keywords: vbaol11.chm678
f1_keywords:
- vbaol11.chm678
ms.prod: outlook
api_name:
- Outlook.SharingItem.InternetCodepage
ms.assetid: a13a44f9-89d1-2839-80e5-de1b8bfab305
ms.date: 06/08/2017
---


# SharingItem.InternetCodepage Property (Outlook)

Returns or sets a  **Long** that determines the Internet code page used by the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **InternetCodepage**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

The Internet code page defines the text encoding scheme used by the item.

The following table lists the values that are supported by the  **InternetCodepage** property.



| **Name**| **Character Set**| **Code Page**|
|Arabic (ISO)|iso-8859-6|28596|
|Arabic (Windows)|windows-1256|1256|
|Baltic (ISO)|iso-8859-4|28594|
|Baltic (Windows)|windows-1257|1257|
|Central European (ISO)|iso-8859-2|28592|
|Central European (Windows)|windows-1250|1250|
|Chinese Simplified (GB2312)|gb2312|936|
|Chinese Simplified (HZ)|hz-gb-2312|52936|
|Chinese Traditional (Big5)|big5|950|
|Cyrillic (ISO)|iso-8859-5|28595|
|Cyrillic (KOI8-R)|koi8-r|20866|
|Cyrillic (KOI8-U)|koi8-u|21866|
|Cyrillic (Windows)|windows-1251|1251|
|Greek (ISO)|iso-8859-7|28597|
|Greek (Windows)|windows-1253|1253|
|Hebrew (ISO-Logical)|iso-8859-8-i|38598|
|Hebrew (Windows)|windows-1255|1255|
|Japanese (EUC)|euc-jp|51932|
|Japanese (JIS)|iso-2022-jp|50220|
|Japanese (JIS-Allow 1 byte Kana)|csISO2022JP|50221|
|Japanese (Shift-JIS)|iso-2022-jp|932|
|Korean|ks_c_5601-1987|949|
|Korean (EUC)|euc-kr|51949|
|Latin 3 (ISO)|iso-8859-3|28593|
|Latin 9 (ISO)|iso-8859-15|28605|
|Thai (Windows)|windows-874|874|
|Turkish (ISO)|iso-8859-9|28599|
|Turkish (Windows)|windows-1254|1254|
|Unicode (UTF-7)|utf-7|65000|
|Unicode (UTF-8)|utf-8|65001|
|US-ASCII|us-ascii|20127|
|Vietnamese (Windows)|windows-1258|1258|
|Western European (ISO)|iso-8859-1|28591|
|Western European (Windows)|windows-1252|1252|
The following table lists the code pages Microsoft recommends that you use for the best compatiblity with older e-mail systems.



| **Name**| **Character Set**| **Code Page**|
|Arabic (Windows)|windows-1256|1256|
|Baltic (ISO)|iso-8859-4|28594|
|Central European (ISO)|iso-8859-2|28592|
|Chinese Simplified (GB2312)|gb2312|936|
|Chinese Traditional (Big5)|big5|950|
|Cyrillic (KOI8-R)|koi8-r|20866|
|Cyrillic (Windows)|windows-1251|1251|
|Greek (ISO)|iso-8859-7|28597|
|Hebrew (Windows)|windows-1255|1255|
|Japanese (JIS)|iso-2022-jp|50220|
|Korean|ks_c_5601-1987|949|
|Thai (Windows)|windows-874|874|
|Turkish (ISO)|iso-8859-9|28599|
|Unicode (UTF-8)|utf-8|65001|
|US-ASCII|us-ascii|20127|
|Vietnamese (Windows)|windows-1258|1258|
|Western European (ISO)|iso-8859-1|28591|

## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

