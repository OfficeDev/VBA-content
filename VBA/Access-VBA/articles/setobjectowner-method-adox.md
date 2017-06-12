---
title: SetObjectOwner method [ADOX]
ms.prod: access
ms.assetid: 22c5d2d9-c7b2-3c3a-0b1f-a2e5bc46395c
ms.date: 06/08/2017
---


# SetObjectOwner method [ADOX]

  

**Applies to:** Access 2013 | Access 2016



Specifies the owner of an object in a  **Catalog**.

## Parameters


-  _ObjectName_
    
- A  **String** value that specifies the name of the object for which to specify the owner.
    
-  _ObjectType_
    
- A  **Long** value which can be one of the **ObjectTypeEnum** constants that specifies the owner type.
    
-  _OwnerName_
    
- A  **String** value that specifies the **Name** of the **User** or **Group** to own the object.
    
-  _ObjectTypeId_
    
- Optional. A  **Variant** value that specifies the GUID for a provider object type not defined by the OLE DB specification. This parameter is required if _ObjectType_ is set to **adPermObjProviderSpecific**; otherwise, it is not used.
    

## Remarks

An error will occur if the provider does not support specifying object owners.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

