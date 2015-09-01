
# Stores.StoreAdd Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when a  ** [Store](1eb22fe9-8849-7476-5388-2515b48591b9.md)** has been added to the current session either programmatically or through user action.

## Syntax

 _expression_. **StoreAdd**( **_Store_**)

 _expression_A variable that represents a  **Stores** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Store|Required| **Store**|The  **Store** to be added to the current session.|

## Remarks

Outlook must be running in order for this event to fire. This event will fire when any of the following occur:


- A store is added through the  **Open Outlook Data File** dialog box, by selecting **Open** and then **Outlook Data File** on the **File** menu.
    
- A store is added through the  **Data Files** tab of the **Account Manager** dialog box.
    
- A store is added successfully by calling the  ** [Namespace.AddStore](c9390982-2408-fda5-a14d-de6f0daaadf1.md)** method.
    


This event will not fire when any of the following occurs:


- When Outlook starts and opens a primary or delegate store. 
    
- If a store is added through the  **Mail** applet in the Microsoft Windows Control Panel and Outlook is not running.
    
- A delegate store is added through the  **Advanced** tab of the **Microsoft Exchange Server** dialog box.
    


You can use this event to determine whether a store has been added and take appropriate actions on items in that store. Otherwise, you would have to resort to polling the  ** [Stores](8915a8e4-9c22-21d5-c492-051d393ce5f7.md)** collection.


## See also


#### Concepts


 [Stores Object](8915a8e4-9c22-21d5-c492-051d393ce5f7.md)
#### Other resources


 [Stores Object Members](f3fec99a-54b2-c13e-d96a-c8c5e2429f99.md)
