---
title: Hyperlink.AddToFavorites Method (Visio)
keywords: vis_sdr.chm15016065
f1_keywords:
- vis_sdr.chm15016065
ms.prod: visio
api_name:
- Visio.Hyperlink.AddToFavorites
ms.assetid: 21a86316-6a59-dc7e-b4f1-0a3d034ba32a
ms.date: 06/08/2017
---


# Hyperlink.AddToFavorites Method (Visio)

Adds a shortcut for a hyperlink address in the presently registered Favorites folder.


## Syntax

 _expression_ . **AddToFavorites**( **_FavoritesTitle_** )

 _expression_ A variable that represents a **Hyperlink** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FavoritesTitle_|Optional| **Variant**|The title to assign to the new shortcut.|

### Return Value

Nothing


## Remarks

If a string is not supplied, the  **AddToFavorites** method uses the hyperlink's **Description** property as the new favorite's title. If the **Description** property is empty, the shortcut is given a generic title, such as Favorite1.

The optional  _favoritesTitle_ argument can specify the full path for the favorites file, for example, "C:\Users\ _username_ \Favorites\My Favorite.URL", or a path relative to the Favorites folder.

From Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA), a call to the  **AddToFavorites** method can take either of these two forms:




```
object.AddToFavorites "SomeString " 
object.AddToFavorites 

```

From C/C++, if a string is supplied, pass a  **Variant** of type VT_BSTR. The application assigns the string as the title of the shortcut. If a string is not supplied, pass a **Variant** of type VT_EMPTY, or of type VT_ERROR and HRESULT DISP_E_PARAMNOTFOUND.


## Example



The following macro shows how to add a hyperlink to a shape and assign a description and address to the hyperlink. Then it shows four ways to use the  **AddToFavorites** method to add the hyperlink to the Favorites folder.



Before running this macro, replace  _address_ with a valid Internet or intranet address, and replace _path_ with a valid path and folder name, including the drive letter, if necessary, on your computer.




```vb
Sub AddToFavorites_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoHyperlink As Visio.Hyperlink 
 
 'Create a new shape to add the hyperlink to. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 Set vsoHyperlink = vsoShape.AddHyperlink 
 
 'Assign a description and an address to the hyperlink. 
 vsoHyperlink.Description = "Web site" 
 vsoHyperlink.Address = "http://address " 
 
 'Use the default name for the new favorite link. 
 vsoHyperlink.AddToFavorites 
 
 'Specify a different name for the new favorite link. 
 'You don't need to specify the URL extension. 
 vsoHyperlink.AddToFavorites "New Favorite Name" 
 
 'Specify a different path to the favorite. 
 vsoHyperlink.AddToFavorites "path\favoriteName " 
 
 'Set a hyperlink base to allow relative hyperlinks. 
 ThisDocument.HyperlinkBase = "path " 
 
 'Specify a relative path to the Favorites folder. 
 'The URL extension is added automatically. 
 vsoHyperlink.AddToFavorites ".\favoriteName " 
 
End Sub
```


