---
title: Documents.AddBlogDocument Method (Word)
keywords: vbawd10.chm158072853
f1_keywords:
- vbawd10.chm158072853
ms.prod: word
api_name:
- Word.Documents.AddBlogDocument
ms.assetid: d47b1b27-a5df-1c82-a8eb-6a4a2853f1ac
ms.date: 06/08/2017
---


# Documents.AddBlogDocument Method (Word)

Returns a  **Document** object that represents a new blog document that Microsoft Word publishes to the account described by the first three parameters.


## Syntax

 _expression_ . **AddBlogDocument**( **_ProviderID_** , **_PostURL_** , **_BlogName_** , **_PostID_** )

 _expression_ An expression that returns a **[Documents](documents-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProviderID_|Required| **String**|A GUID that is the unique value a provider uses when they register themselves with Word.|
| _PostURL_|Required| **String**|The URL that is used to add posts to the blog.|
| _BlogName_|Required| **String**|A display name for the blog that will be used in Word.|
| _PostID_|Optional| **String**|The ID for an existing post with which to populate the document created by using the  **AddBlogDocument** method.|

## Remarks

This method creates a new document, and it also registers the specified blog account with Word if it is not already registered. In addition, if the PostID parameter is specified, the new document is populated with the contents of the post specified by the value of the PostID parameter, from the provider's Web site.


## See also


#### Concepts


[Documents Collection Object](documents-object-word.md)

