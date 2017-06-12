---
title: Application.ChangeFileOpenDirectory Method (Publisher)
keywords: vbapb10.chm131124
f1_keywords:
- vbapb10.chm131124
ms.prod: publisher
api_name:
- Publisher.Application.ChangeFileOpenDirectory
ms.assetid: 9178881c-2f7f-9063-31d1-14d4745f0666
ms.date: 06/08/2017
---


# Application.ChangeFileOpenDirectory Method (Publisher)

Sets the folder in which Microsoft Publisher searches for documents. The specified folder's contents are listed the next time the  **Open Publication** dialog box ( **File** menu) is displayed.


## Syntax

 _expression_. **ChangeFileOpenDirectory**( **_Dir_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Dir|Required| **String**|The directory path.|

## Remarks

Publisher searches the specified folder for documents until the user changes the folder in the  **Open Publication** dialog box or the current Publisher session ends. Use the **[PathForPublications](options-pathforpublications-property-publisher.md)** property of the  **Options** object to change the default folder for documents in every Publisher session.


## Example

This example changes the folder in which Publisher searches for documents. (Note that PathToDirectory must be replaced with a valid file path for this example to work.)


```vb
Sub ChangeOpenPath() 
 ChangeFileOpenDirectory Dir:="PathToDirectory" 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

