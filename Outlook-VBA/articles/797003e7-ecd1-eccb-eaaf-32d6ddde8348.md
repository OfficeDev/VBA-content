
# Application Object (Outlook)

 **Last modified:** July 28, 2015

Represents the entire Microsoft Outlook application.

## Remarks

 This is the only object in the hierarchy that can be returned by using the ** [CreateObject](09b6ff5b-a750-c07d-7499-c1f8a00214fe.md)**method or the intrinsic Visual Basic  **GetObject** function.

The Outlook  **Application** object has several purposes:


- As the root object, it allows access to other objects in the Outlook hierarchy.
    
- It allows direct access to a new item created by using  ** [CreateItem](e5fbf367-db16-5042-823e-68e6b805e612.md)**, without having to traverse the object hierarchy.
    
- It allows access to the active interface objects (the explorer and the inspector).
    
When you use Automation to control Outlook from another application, you use the  **CreateObject** method to create an Outlook **Application** object.


## Example

The following Visual Basic for Applications (VBA) example starts Outlook (if it's not already running) and opens the default Inbox folder.


```
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myFolder= _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
myFolder.Display
```

The following Visual Basic for Applications (VBA) example uses the  **Application** object to create and open a new contact.




```
Set myItem = Application.CreateItem(olContactItem) 
 
myItem.Display
```


## See also


#### Concepts


 [Outlook Object Model Reference](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)
#### Other resources


 [Application Object Members](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)
