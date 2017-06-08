---
title: Templates Object (Word)
keywords: vbawd10.chm2466
f1_keywords:
- vbawd10.chm2466
ms.prod: word
ms.assetid: de62f768-011a-7446-48c3-1c4512da5f7c
ms.date: 06/08/2017
---


# Templates Object (Word)

A collection of  **[Template](template-object-word.md)** objects that represent all the templates that are currently available. This collection includes open templates, templates attached to open documents, and global templates loaded in the **Templates and Add-ins** dialog box.


## Remarks

Use the  **Templates** property to return the **Templates** collection. The following example displays the path and file name of each template in the **Templates** collection.


```vb
For Each aTemp In Templates 
 MsgBox aTemp.FullName 
Next aTemp
```

The  **Add** method isn't available for the **Templates** collection. Instead, you can add a template to the **Templates** collection by doing any of the following:


- Using the  **Open** method with the **Documents** collection to open a document based on a template or a template
    
- Using the  **Add** method with the **Documents** collection to open a new document based on a template
    
- Using the  **Add** method with the **Addins** collection to load a global template
    
- Using the  **AttachedTemplate** property with the **Document** object to attach a template to a document
    
Use  **Templates** (Index), where Index is the template name or the index number, to return a single **Template** object. The following example saves the Dot1.dot template.




```
Templates("C:\MSOffice\WinWord\Templates\Dot1.dot").Save
```

The index number represents the position of the template in the  **Templates** collection. The following example displays the file name of the first template in the **Templates** collection.




```vb
MsgBox Templates(1).FullName
```

Use the  **NormalTemplate** property to return a template object that refers to the Normal template. Use the **AttachedTemplate** property to return the template attached to the specified document.

Use the  **DefaultFilePath** property to determine the location of user or workgroup templates (that is, the folder where you want to store these templates). The following example displays the user template folder from the **File Locations** tab in the **Options** dialog box.




```vb
MsgBox Options.DefaultFilePath(wdUserTemplatePath)
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

