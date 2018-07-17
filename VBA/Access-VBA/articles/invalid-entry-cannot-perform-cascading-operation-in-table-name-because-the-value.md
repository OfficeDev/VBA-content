---
title: Invalid entry. Cannot perform cascading operation in table <name> because the value entered is too large for field <name>. (Error 3411)
ms.prod: access
ms.assetid: 286a606c-72c0-7dab-0dc7-0fba19d683bb
ms.date: 06/08/2017
---


# Invalid entry. Cannot perform cascading operation in table <name> because the value entered is too large for field <name>. (Error 3411)

  

**Applies to:** Access 2013 | Access 2016

You are trying to save a value to a primary key field that is included in a relationship.

In Microsoft Access, the  **Cascade Update Related Fields** option is selected for the relationship; or, in DAO code, the **dbRelationUpdateCascade** option is specified for the **Relation** object's **Attributes** property. Setting either one of these options will cause your application to attempt to update the matching field in the related table.
To save your changes to this field, the text you enter must be no longer than the field size of the related field that your application is trying to update for you. In this case, the related field has a shorter field size than the field you are updating. To avoid this problem in the future, set the  **Size** property for both of the matching fields to the same value.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

