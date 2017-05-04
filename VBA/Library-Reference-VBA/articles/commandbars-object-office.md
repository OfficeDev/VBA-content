---
title: CommandBars Object (Office)
keywords: vbaof11.chm242000
f1_keywords:
- vbaof11.chm242000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBars
ms.assetid: 0e312e21-14ee-5055-d604-b66e61c53b47
---


# CommandBars Object (Office)

A collection of  **CommandBar** objects that represent the command bars in the container application.


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use the  **CommandBars** property to return the **CommandBars** collection. The following example displays in the **Immediate** window both the name and local name of each menu bar and toolbar, and it displays a value that indicates whether the menu bar or toolbar is visible.


```
For Each cbar in CommandBars 
    Debug.Print cbar.Name, cbar.NameLocal, cbar.Visible 
Next
```

Use the  **Add** method to add a new command bar to the collection. The following example creates a custom toolbar named "Custom1" and displays it as a floating toolbar.




```
Set cbar1 = CommandBars.Add(Name:="Custom1", Position:=msoBarFloating) 
cbar1.Visible = True
```

Use enumName, where  _index_ is the name or index number of a command bar, to return a single **CommandBar** object. The following example docks the toolbar named "Custom1" at the bottom of the application window.




```
CommandBars("Custom1").Position = msoBarBottom
```


 **Note**  You can use the name or index number to specify a menu bar or toolbar in the list of available menu bars and toolbars in the container application. However, you must use the name to specify a menu, shortcut menu, or submenu (all of which are represented by  **CommandBar** objects). If two or more custom menus or submenus have the same name, enumName returns the first one. To ensure that you return the correct menu or submenu, locate the pop-up control that displays that menu. Then apply the **CommandBar** property to the pop-up control to return the command bar that represents that menu.


## Events



|**Name**|
|:-----|
|[OnUpdate](http://msdn.microsoft.com/library/4da9354b-92ed-d85e-f667-c01dfec07689%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/544cfa94-924a-90ca-d716-c7b2f9e8732f%28Office.15%29.aspx)|
|[CommitRenderingTransaction](http://msdn.microsoft.com/library/a3174734-305b-03dc-2da1-3d25fd74185d%28Office.15%29.aspx)|
|[ExecuteMso](http://msdn.microsoft.com/library/6f608475-7a79-48c7-abff-86d9ab07fe80%28Office.15%29.aspx)|
|[FindControl](http://msdn.microsoft.com/library/07ec0c01-3cf4-3165-cfb2-c596b5e39abd%28Office.15%29.aspx)|
|[FindControls](http://msdn.microsoft.com/library/79c46884-816d-def6-2bff-85b59b0831ea%28Office.15%29.aspx)|
|[GetEnabledMso](http://msdn.microsoft.com/library/68af6404-53ee-4c69-51fa-4d489736d228%28Office.15%29.aspx)|
|[GetImageMso](http://msdn.microsoft.com/library/36261e2b-9cbf-b0b6-5892-63bbb2f93959%28Office.15%29.aspx)|
|[GetLabelMso](http://msdn.microsoft.com/library/1ab6f700-e3c3-a89d-790f-10c27a6b495c%28Office.15%29.aspx)|
|[GetPressedMso](http://msdn.microsoft.com/library/97811bb6-cc5c-eccc-9149-76bdfa37541f%28Office.15%29.aspx)|
|[GetScreentipMso](http://msdn.microsoft.com/library/23411622-2b35-0c0e-9373-9bc75c5e433e%28Office.15%29.aspx)|
|[GetSupertipMso](http://msdn.microsoft.com/library/e116402f-bbb7-8cd3-6305-7daf85feb514%28Office.15%29.aspx)|
|[GetVisibleMso](http://msdn.microsoft.com/library/ab916050-e1af-0752-9734-23d0fe27542f%28Office.15%29.aspx)|
|[ReleaseFocus](http://msdn.microsoft.com/library/2ddca1e1-b8f4-a09c-120d-498b816747c4%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionControl](http://msdn.microsoft.com/library/70097691-a771-4f7d-020b-2a9d33e18fa0%28Office.15%29.aspx)|
|[ActiveMenuBar](http://msdn.microsoft.com/library/8f341f53-418c-6d05-ac0b-e45a6b2baa0d%28Office.15%29.aspx)|
|[AdaptiveMenus](http://msdn.microsoft.com/library/1b8c1a2a-9fe1-4148-6e03-5bf48f137d6f%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/98ce76f8-c2ef-0304-97c6-70e2567700e7%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/10b19483-f9a0-dd0d-512f-74afc1ddfe8b%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/7841f7b3-2ae7-9264-37e7-c359d583a2a1%28Office.15%29.aspx)|
|[DisableAskAQuestionDropdown](http://msdn.microsoft.com/library/a0954aa4-256c-4a14-6bab-959a00e9367d%28Office.15%29.aspx)|
|[DisableCustomize](http://msdn.microsoft.com/library/cbebdaa7-2e8d-af73-fd18-03b3b11f98ac%28Office.15%29.aspx)|
|[DisplayFonts](http://msdn.microsoft.com/library/25a9ede7-3575-6706-406d-a5b656cd965e%28Office.15%29.aspx)|
|[DisplayKeysInTooltips](http://msdn.microsoft.com/library/de132c5f-bc9f-c335-28ff-b9459c912b2c%28Office.15%29.aspx)|
|[DisplayTooltips](http://msdn.microsoft.com/library/98b62729-d1c8-a6dc-328e-8dbb6bbd80dc%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/bca38d83-67cb-2cba-ddfa-918a5b2ff508%28Office.15%29.aspx)|
|[LargeButtons](http://msdn.microsoft.com/library/bcacab92-9779-5061-f68a-69722210e14e%28Office.15%29.aspx)|
|[MenuAnimationStyle](http://msdn.microsoft.com/library/bd79a55a-23f4-6056-649b-9dc384b597aa%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/7819df1a-1f63-793c-54f3-c1129fd6cdff%28Office.15%29.aspx)|

## See also


#### Other resources


[CommandBars Object Members](http://msdn.microsoft.com/library/c11db22d-b7bb-20a2-a455-e441cb8d5bc0%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
