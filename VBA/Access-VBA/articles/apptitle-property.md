---
title: AppTitle Property
keywords: vbaac10.chm5187013
f1_keywords:
- vbaac10.chm5187013
ms.prod: access
api_name:
- Access.AppTitle
ms.assetid: a505f465-7813-6677-dd80-21a757c9d422
ms.date: 06/08/2017
---


# AppTitle Property

  

**Applies to:** Access 2013 | Access 2016

You can use the  **AppTitle** property to specify the text that appears in the application database's title bar. For example, you can use the **AppTitle** property to specify that the string "Inventory Control" appear in the title bar of your Inventory Control database application.


## Setting

The  **AppTitle** property is a string expression containing the text to appear in the title bar.

The easiest way to set this property is by using the  **Application Title** option in the **Access Options** dialog box, available by clicking the click the **Microsoft Office Button**
![File menu button](images/O12FileMenuButton_ZA10077102.gif) and then clicking the **Current Database** category. You can also set this property by using a macro or Visual Basic .

To set the  **AppTitle** property by using a macro or Visual Basic, you must first either set the property in the **Access Options** dialog box once or create the property in the following ways:


- In a Microsoft Access database, you can add it by using the  **CreateProperty** method and append it to the **Properties** collection of the **Database** object.
    
- In a Microsoft Access project (.adp), you can add it to the  **AccessObjectProperties** collection of the **CurrentProject** object by using the **Add** method.
    
You must also use the RefreshTitleBar method to make any changes visible immediately.


## Remarks

If this property isn't set, the string "Microsoft Access" appears in the title bar.

This property's setting takes effect immediately after setting the property in code (as long as the code includes the  **RefreshTitleBar** method) or closing the **Access Options** dialog box.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

