---
title: IBlogPictureExtensibility.PublishPicture Method (Office)
keywords: vbaof11.chm329003
f1_keywords:
- vbaof11.chm329003
ms.prod: office
api_name:
- Office.IBlogPictureExtensibility.PublishPicture
ms.assetid: b8adbff6-a446-047d-59cd-359e69960d22
ms.date: 06/08/2017
---


# IBlogPictureExtensibility.PublishPicture Method (Office)

Used to post a picture object to its final destination in a blog.


## Syntax

 _expression_. **PublishPicture**( **_Account_**, **_ParentWindow_**, **_Document_**, **_userName_**, **_Password_**, **_Image_**, **_PictureURI_** )

 _expression_ An expression that returns a **IBlogPictureExtensibility** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents user's password stored in the registry account settings.|
| _Image_|Required|**Unknown**|Represents the name of the image file.|
| _PictureURI_|Required|**String**|The URI of the picture.|

## Remarks

This method is called during xHTML generation and hands off a .PNG representation of the picture. The location where it was published is then returned which is then added to an IMG tag in the xHTML output.


## See also


#### Concepts


[IBlogPictureExtensibility Object](iblogpictureextensibility-object-office.md)
#### Other resources


[IBlogPictureExtensibility Object Members](iblogpictureextensibility-members-office.md)

