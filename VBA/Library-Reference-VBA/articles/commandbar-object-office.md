---
title: CommandBar Object (Office)
keywords: vbaof11.chm3000
f1_keywords:
- vbaof11.chm3000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBar
ms.assetid: 78603954-40aa-64cb-c407-2e0820d65231
---


# CommandBar Object (Office)

Represents a command bar in the container application. The  **CommandBar** object is a member of the **CommandBars** collection.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use  **CommandBars** ( _index_ ), where _index_ is the name or index number of a command bar, to return a single **CommandBar** object. The following example steps through the collection of command bars to find the command bar named "Forms." If it finds this command bar, the example makes it visible and protects its docking state. In this example, the variable **cb** represents a **CommandBar** object.


```
foundFlag = False  
For Each cb In CommandBars 
    If cb.Name = "Forms" Then 
        cb.Protection = msoBarNoChangeDock 
        cb.Visible = True  
        foundFlag = True  
    End If 
Next cb 
If Not foundFlag Then 
    MsgBox "The collection does not contain a Forms command bar." 
End If
```

You can use a name or index number to specify a menu bar or toolbar in the list of available menu bars and toolbars in the container application. However, you must use a name to specify a menu, shortcut menu, or submenu (all of which are represented by  **CommandBar** objects). This example adds a new menu item to the bottom of the **Tools** menu. When clicked, the new menu item runs the procedure named "qtrReport."




```
Set newItem = CommandBars("Tools").Controls.Add(Type:=msoControlButton) 
With newItem 
    .BeginGroup = True  
    .Caption = "Make Report" 
    .FaceID = 0 
    .OnAction = "qtrReport" 
End With
```

If two or more custom menus or submenus have the same name,  **CommandBars(index)** returns the first one. To ensure that you return the correct menu or submenu, locate the pop-up control that displays that menu. Then apply the **CommandBar** property to the pop-up control to return the command bar that represents that menu. Assuming that the third control on the toolbar named "Custom Tools" is a pop-up control, this example adds the **Save** command to the bottom of that menu.




```
Set viewMenu = CommandBars("Custom Tools").Controls(3) 
viewMenu.Controls.Add ID:=3    'ID of Save command is 3
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/6976f273-dbd4-5f3d-52ef-0d6d5cc886c9%28Office.15%29.aspx)|
|[FindControl](http://msdn.microsoft.com/library/d5ff45de-a356-0dab-4233-88326d08535a%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/96dfb3cc-a53c-ea7f-eb98-96a983faa681%28Office.15%29.aspx)|
|[ShowPopup](http://msdn.microsoft.com/library/e501b7d2-2606-976c-b391-1aa8fa07f105%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AdaptiveMenu](http://msdn.microsoft.com/library/1e6920bb-af66-951c-e689-399d9cf5d662%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/afe6da31-95af-1a41-4ce7-f5b0c4d65ad7%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/f7e4c581-2019-9fca-5e9e-15db4d656269%28Office.15%29.aspx)|
|[Context](http://msdn.microsoft.com/library/e7b8a7e5-0799-84e8-c7e3-5f713971099d%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/5c025bc5-9266-18a2-21ee-6aee478fb322%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/7de5e3d3-9a02-536f-1afb-58afe017cd44%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/4a332d30-4aa9-1355-2d26-0d4f0529d488%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/9a5c84ae-29c0-0ff3-74f4-864c978336d2%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/a8b2e075-4c2a-5f53-a343-579e7e244c8f%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/2353aef6-aaa1-76b9-33da-57bbe1df30af%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/4d578782-b59d-3dd7-be99-b9d79f8f3eaa%28Office.15%29.aspx)|
|[NameLocal](http://msdn.microsoft.com/library/3afad045-aaf8-8775-574e-faaccde7d270%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/6b9e1f55-84a6-afa0-a18d-3e2d7a10b2f5%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/b1e80bc0-1586-523b-a9ec-70c76fa54252%28Office.15%29.aspx)|
|[Protection](http://msdn.microsoft.com/library/59f9e9d3-251c-93a6-fa49-75fa7c4f6659%28Office.15%29.aspx)|
|[RowIndex](http://msdn.microsoft.com/library/6dd5576c-0a46-9a72-9c4e-fcf685097b77%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/1bac668a-0caa-d185-cc07-ba55809c79fe%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/e023edd9-a8f4-c20f-c6b1-c434182bd748%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/c7057c83-ea8d-c167-a650-d784d5e6dd1f%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/ae092193-59fd-25a1-c1d0-ebe6d6532756%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[CommandBar Object Members](http://msdn.microsoft.com/library/e3756e7e-56a8-33a4-722f-640e5cc69b6d%28Office.15%29.aspx)
