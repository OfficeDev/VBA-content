
# NavigationFolder Object (Outlook)

 **Last modified:** July 28, 2015

Represents a navigation folder displayed in a navigation group of a navigation module in the Navigation Pane.

## Remarks

Use the  ** [Item](1688b2ef-a4a1-fc8a-513e-0d5e234f10dd.md)** method to retrieve a **NavigationFolder** object from the ** [NavigationFolders](ecff93b8-0c3f-5f31-5b61-c46d2622d2af.md)** collection of the parent ** [NavigationGroup](a96eb2b1-af1f-71b2-6a0b-dcb5078beb1f.md)** object. Use the ** [Add](f88fd69a-8684-bfc4-bc20-1cff5c44974e.md)** method of the **NavigationFolders** collection to create a new **NavigationFolder** object based on an existing ** [Folder](3cf6cda8-6d70-666e-2643-9d9c5b9cacfc.md)** object.

Use the  ** [Folder](0d8edd40-3f8d-dc2b-5cba-80ed1662cc48.md)** method to return or set the **Folder** object on which the **NavigationFolder** object is based.

Use the  ** [IsSelected](a8fb9430-0477-2417-0dba-e30e9f8ebe8d.md)** property to determine if the navigation folder is selected and the ** [Position](cfa86104-c191-51f8-4da3-dc3c26d6a7ed.md)** property to return or set the display position of the navigation folder within the Navigation Pane. You can also use the ** [DisplayName](51bdcbaf-0fa7-8cba-953d-13da4a5abc27.md)** property to return the display name of the navigation folder within the Navigation Pane.

Use the  ** [IsRemovable](9fff5f32-2ac4-5ed3-c6d5-10962de8b34f.md)** property to determine if a navigation folder can be removed from the **NavigationFolders** collection and the ** [IsSideBySide](00a49ce6-ad74-1f24-2aaa-e79a3409c9c9.md)** property to return or set the viewing mode for a navigation folder associated with a ** [CalendarModule](9203024d-9cef-75e0-600f-f3899e24761a.md)** object.


## See also


#### Concepts


 [Outlook Object Model Reference](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)
#### Other resources


 [NavigationFolder Object Members](1ec2e16d-c7ca-86b1-9283-839a2b9aca05.md)
