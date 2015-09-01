
# Store Data in a StorageItem for a Solution

 **Last modified:** July 28, 2015

 _**Applies to:** Outlook 2013_

This topic describes how to store private application data in solution storage provided by the Outlook object model.


1. Determine the folder where you would like to store your application data. 
    
     **Note**  Because solution storage is created as hidden items in a folder, you can only store solution data if the store provider supports hidden items and the client has rights to write to that folder.
2. Use  ** [Folder.GetStorage](cc5ee63b-7d11-6340-8392-8b35a689a28c.md)** to obtain either an existing ** [StorageItem](41776bc3-b838-2755-fd6b-3b5012fb9ae5.md)** object or a new **StorageItem** object if one does not already exist.
    
3. Use  ** [StorageItem.Size](7bf2fd39-8705-aa1b-af76-a3a21073d152.md)** to determine if the **StorageItem** is new. If it is, then use the ** [Add](88b86622-2234-77be-41e7-b76b0b3a75ad.md)** method of ** [StorageItem.UserProperties](0a08e77c-1665-a612-2f47-ef1c3fc331d2.md)** to create a custom property **Order Number**.
    
4. Set the  **Order Number** property. This assumes that an existing **StorageItem** already has the custom property **Order Number** defined.
    
5. Use  ** [StorageItem.Save](9462a342-294a-175e-7e8f-d416f0959f69.md)** to save the **StorageItem** object as a hidden item in the folder.
    



```
Sub StoreData() 
 Dim oInbox As Folder 
 Dim myStorage As StorageItem 
 Dim myPrivateProperty As UserProperty 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 ' Get an existing instance of StorageItem by subject, or create new if it doesn't exist 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 If myStorage.Size = 0 Then 
 'There was no existing StorageItem by this subject, so created a new one 
 'Create a custom property for Order Number 
 Set myPrivateProperty = myStorage.UserProperties.Add("Order Number", olNumber) 
 Else 
 'Assume that existing storage has the Order Number property already 
 Set myPrivateProperty = myStorage.UserProperties("Order Number") 
 End If 
 myPrivateProperty.Value = lngOrderNumber 
 myStorage.Save 
End Sub
```

