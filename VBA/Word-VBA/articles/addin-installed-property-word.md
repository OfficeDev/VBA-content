---
title: AddIn.Installed Property (Word)
keywords: vbawd10.chm159252484
f1_keywords:
- vbawd10.chm159252484
ms.prod: word
api_name:
- Word.AddIn.Installed
ms.assetid: 5bca123c-c75f-23f0-15d1-cf9f662de8da
ms.date: 06/08/2017
---


# AddIn.Installed Property (Word)

 **True** if the specified add-in is installed (loaded). Add-ins that are loaded are selected in the **Templates and Add-ins** dialog box. Read/write **Boolean** .


## Syntax

 _expression_ . **Installed**

 _expression_ An expression that returns an **[AddIn](addin-object-word.md)** object.


## Remarks

Uninstalled add-ins are included in the  **[AddIns](addins-object-word.md)** collection. To remove a template or WLL from the AddIns collection, apply the **[Delete](addin-delete-method-word.md)** method to the **AddIn** object (the add-in name is removed from the **Templates and Add-ins** dialog box). To unload all templates and WLLs, apply the **[Unload](addins-unload-method-word.md)** method to the **AddIns** collection.


## Example

This example unloads the global template named "Gallery.dot."


```vb
Addins("Gallery.dot").Installed = False
```

This example loads FindAll.wll.




```vb
Addins("C:\Templates\FindAll.wll").Installed = True
```

This example loads Custom.dot.




```vb
AddIns("C:\Program Files\Microsoft Office\" _ 
 &; "Templates\Custom.dot").Installed = True
```

This example displays a message on the status bar if Dot1.dot is loaded as a global template.




```vb
If AddIns("Dot1.dot").Installed = True Then _ 
 StatusBar = "Dot1.dot is loaded"
```


## See also


#### Concepts


[AddIn Object](addin-object-word.md)

