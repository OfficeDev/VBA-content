---
title: Size Property (ADO Stream)
ms.prod: access
ms.assetid: deb84313-36d1-fa49-e4cd-daecab96f343
ms.date: 06/08/2017
---


# Size Property (ADO Stream)

  

**Applies to:** Access 2013 | Access 2016



Indicates the size of the stream in number of bytes.

## Return Values

Returns a  **Long** value that specifies the size of the stream in number of bytes. The default value is the size of the stream, or -1 if the size of the stream is not known.


## Remarks

 **Size** can be used only with open[Stream](http://msdn.microsoft.com/library/d49b1514-e0b4-0aca-d5c2-8266f3f4fe65%28Office.15%29.aspx) objects.


 **Note**  Any number of bits can be stored in a  **Stream** object, limited only by system resources. If the **Stream** contains more bits than can be represented by a **Long** value, **Size** is truncated and therefore does not accurately represent the length of the **Stream**.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

