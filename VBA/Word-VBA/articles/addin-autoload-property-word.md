---
title: AddIn.Autoload Property (Word)
keywords: vbawd10.chm159252486
f1_keywords:
- vbawd10.chm159252486
ms.prod: word
api_name:
- Word.AddIn.Autoload
ms.assetid: 320b5624-2b00-991c-18ac-568c87caff42
ms.date: 06/08/2017
---


# AddIn.Autoload Property (Word)

 **True** if the specified add-in is automatically loaded when Word is started. Add-ins located in the Startup folder in the Word program folder are automatically loaded. Read-only **Boolean** .


## Syntax

 _expression_ . **Autoload**

 _expression_ A variable that represents a **[AddIn](addin-object-word.md)** object.


## Example

This example displays the name of each add-in that is automatically loaded when Word is started.


```vb
Dim addinLoop as AddIn 
Dim blnFound as Boolean 
 
blnFound = False 
 
For Each addinLoop In AddIns 
 With addinLoop 
 If .Autoload = True Then 
 MsgBox .Name 
 blnFound = True 
 End If 
 End With 
Next addinLoop 
 
If blnFound <> True Then _ 
 MsgBox "No add-ins were loaded automatically."
```

This example determines whether the add-in named "Gallery.dot" was automatically loaded.




```vb
Dim addinLoop as AddIn 
 
For Each addinLoop In AddIns 
 If InStr(LCase$(addinLoop.Name), "gallery.dot") > 0 Then 
 If addinLoop.Autoload = True Then Msgbox "Autoload" 
 End If 
Next addinLoop
```


## See also


#### Concepts


[AddIn Object](addin-object-word.md)

