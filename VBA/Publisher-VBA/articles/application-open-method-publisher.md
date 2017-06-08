---
title: Application.Open Method (Publisher)
keywords: vbapb10.chm131128
f1_keywords:
- vbapb10.chm131128
ms.prod: publisher
api_name:
- Publisher.Application.Open
ms.assetid: 560ac406-f058-8fd8-4b6d-978ff369de9b
ms.date: 06/08/2017
---


# Application.Open Method (Publisher)

Returns a  **[Document](document-object-publisher.md)** object that represents the newly opened publication.


## Syntax

 _expression_. **Open**( **_Filename_**,  **_ReadOnly_**,  **_AddToRecentFiles_**,  **_SaveChanges_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Required| **String**|The name of the publication (paths are accepted).|
|ReadOnly|Optional| **Boolean**| **True** to open the publication as read-only. Default is **False**.|
|AddToRecentFiles|Optional| **Boolean**| **True** (default) to add the file name to the list of recently used files at the bottom of the File menu.|
|SaveChanges|Optional| **PbSaveOptions**|Specifies what Microsoft Publisher should do if there is already an open publication with unsaved changes.|
|OpenConflictDocument|Optional| **Boolean**| **True** to open the local conflict publication if there is an offline conflict. Default is **False**.|

### Return Value

Document


## Remarks

Because Publisher has a single document interface, the  **Open** method works only when you open a new instance of Publisher. The code sample below shows how to create a new, visible instance of Publisher. When finished with the second instance, you can set the application window's [Visible](window-visible-property-publisher.md)property to  **False**, but the process continues to run in the background, even though it is not visible. To close the second instance, you must set the object equal to  **Nothing**.

The SaveChanges parameter can be one of the  **PbSaveOption** constants declared in the Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbDoNotSaveChanges**|Close the open publication without saving any changes. |
| **pbPromptToSaveChanges**|Prompt the user whether to save changes in the open publication. The default.|
| **pbSaveChanges**|Save the open publication before closing it.|

## Example

This example creates a second instance of Publisher and opens the specified publication as read-only. 

For this example to work, you must replace  _PathToFile_ with the path to an existing publication.




```vb
Sub OpenNewPub() 
 Dim appPub As New Publisher.Application 
 appPub.Open FileName:="PathToFile", _ 
 ReadOnly:=True, AddToRecentFiles:=False, _ 
 SaveChanges:=pbPromptToSaveChanges 
 appPub.ActiveWindow.Visible = True 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

