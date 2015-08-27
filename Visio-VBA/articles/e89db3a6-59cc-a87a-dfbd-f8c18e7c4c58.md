
# Menu.Caption Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets or sets the caption for a menu. Read/write.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Caption**

 _expression_A variable that represents a  **Menu** object.


### Return Value

String


## Remarks
<a name="sectionSection1"> </a>


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.




- Use an ampersand (&amp;) in the string to cause the next character in the string to become the shortcut key for that menu or menu item. For example, the string "F _&amp;o_rmat" causes  _o_ to become the shortcut key for that menu item in that one menu.
    
- Use "" in the string to display a double quotation mark on the menu.
    
- Use &amp;&amp; in the string to display an ampersand on the menu.
    



## Example
<a name="sectionSection2"> </a>

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Caption** property. It adds a menu and menu item to the **Add-ins** tab and sets the menu and menu item's **Caption** properties.

To restore the built-in user interface in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```
 
Public Sub Caption_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoMenuSets As Visio.MenuSets 
 Dim vsoMenuSet As Visio.MenuSet 
 Dim vsoMenus As Visio.Menus 
 Dim vsoMenu As Visio.Menu 
 Dim vsoMenuItems As Visio.MenuItems 
 Dim vsoMenuItem As Visio.MenuItem 
 
 'Get a UIObject object that represents Microsoft Visio built-in menus. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Get the MenuSets collection. 
 Set vsoMenuSets = vsoUIObject.MenuSets 
 
 'Get the drawing window menu set. 
 Set vsoMenuSet = vsoMenuSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the Menus collection. 
 Set vsoMenus = vsoMenuSet.Menus 
 
 'Add a new menu before the Window menu. 
 Set vsoMenu = vsoMenus.AddAt(7) 
 vsoMenu.Caption = "MyNewMenu" 
 
 'Get the MenuItems collection. 
 Set vsoMenuItems = vsoMenu.MenuItems 
 
 'Add a menu item to the new menu. 
 Set vsoMenuItem = vsoMenuItems.Add 
 
 'Set the Caption property for the new menu item. 
 vsoMenuItem.Caption = "&amp;MyNewMenuItem" 
 
 'Tell Visio to use the new UI when the document is active. 
 ThisDocument.SetCustomMenus vsoUIObject 
 
End Sub 

```

