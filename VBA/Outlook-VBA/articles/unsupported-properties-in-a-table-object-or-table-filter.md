---
title: Unsupported Properties in a Table Object or Table Filter
ms.prod: outlook
ms.assetid: 0e37f03f-7677-ca29-d0b2-8b45c026e5f1
ms.date: 06/08/2017
---


# Unsupported Properties in a Table Object or Table Filter

This topic lists the properties that you cannot add to a  **[Table](table-object-outlook.md)** or use in a **Table** filter. You cannot add these properties through **[Columns.Add](columns-add-method-outlook.md)**, and you cannot specify these properties in a filter used by the following methods:


-  **[Folder.GetTable](folder-gettable-method-outlook.md)**
    
-  **[Search.GetTable](search-gettable-method-outlook.md)** (Note that the filter is derived from the **[Search](search-object-outlook.md)** object returned by **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)**)
    
-  **[Table.FindRow](table-findrow-method-outlook.md)**
    
-  **[Table.Restrict](table-restrict-method-outlook.md)**
    

| **Properties**| **In Table Object**| **In Table Filter**| **Comments**|
|:-----|:-----|:-----|:-----|
|Binary properties|Supported |Not supported|If you add a binary property to a  **Table** referencing its namespace, the value of the property in the **Table** is in binary. You can use **[Row.BinaryToString](row-binarytostring-method-outlook.md)** to convert the value to a string.|
|Body properties, including  **Body**,  **HTMLBody**, **http://schemas.microsoft.com/mapi/proptag/0x10130102**<br> (for **PidTagHtml**), and **http://schemas.microsoft.com/mapi/proptag/0x10090102** (for **PidTagRtfCompressed**)|The  **Body** property is supported with a condition that only the first 255 bytes of the value are stored in a **Table**. Other properties representing the body content in HTML or RTF are not supported. <br> Because only the first 255 bytes of  **Body** is stored in a **Table**, if you want to obtain the full body content of an item in text or HTML, use the item's  **EntryID** in **[GetItemFromID](namespace-getitemfromid-method-outlook.md)** to obtain the item object. Then retrieve the full value of **Body** through the item object.|Only the  **Body** property represented in text is supported in a filter. This means that the property must be referenced in a DASL filter as **urn:schemas:httpmail:textdescription**, and you cannot filter on any HTML tags in the body. To improve performance, use context indexer keywords in the filter to match strings in the body.||
|Computed properties, such as  **AutoResolvedWinner** and **BodyFormat**. See below for a complete list of computed properties.|Not supported|Not supported|To obtain the value of a computed property for an item in a  **Table**, use the item's  **EntryID** in **GetItemFromID** to obtain the item object. Then retrieve the property value through the item object.|
|Multi-valued properties, such as  **Categories**,  **[Children](contactitem-children-property-outlook.md)**,  **[Companies](contactitem-companies-property-outlook.md)**, and  **[VotingOptions](mailitem-votingoptions-property-outlook.md)**|Supported|Although both Jet and DASL filters both support multi-valued properties, use content indexing in DASL filters for more efficient filtering. For more information, see  [Filtering Items Using a Comparison with a Keywords Property](filtering-items-using-a-comparison-with-a-keywords-property.md).|The format of the values of a multi-valued property in a  **Table** depends on whether the property was added with its explicit built-in name or with a name referencing its namespace. If the property is added with its explicit built-in name, the value in the **Table** is a comma-delimited string. Otherwise, the value is a variant array. For more information, see [How to: Access the Values of a Multi-valued Property in a Table](access-the-values-of-a-multi-valued-property-in-a-table.md).|
|Properties returning an object, such as  **Attachments**,  **Parent**,  **Recipients**,  **RecurrencePattern**, and  **UserProperties**.|Not supported if property is referenced by its explicit built-in name; supported if property is referenced by its namespace.|Not supported if property is expressed in a Jet query; supported if property is expressed in a DASL query.||


## Unsupported Computed Properties

If you attempt to add one of the computed properties listed below using  **Columns.Add**, referencing the property either by the explicit property name or by namespace, you will get the error,  **IDS_ERR_BLOCKED_PROPERTY**. To determine the value of these properties, obtain the item object using its Entry ID and then use the item object to determine the property value (as in  `object.property`):


-  **AutoResolvedWinner**
    
-  **BodyFormat**
    
-  **Class**
    
-  **ContactNames**
    
-  **Companies**
    
-  **[DLName](distlistitem-dlname-property-outlook.md)**
    
-  **DownloadState**
    
-  **FlagIcon**
    
-  **HtmlBody**
    
-  **InternetCodePage**
    
-  **IsConflict**
    
-  **IsMarkedAsTask**
    
-  **MeetingWorkspaceURL**
    
-  **MemberCount**
    
-  **[Permission](mailitem-permission-property-outlook.md)**
    
-  **[PermissionService](mailitem-permissionservice-property-outlook.md)**
    
-  **[RecurrenceState](appointmentitem-recurrencestate-property-outlook.md)**
    
-  **[ResponseState](taskitem-responsestate-property-outlook.md)**
    
-  **Saved**
    
-  **Sent**
    
-  **Submitted**
    
-  **TaskSubject**
    
-  **Unread**
    
-  **[VotingOptions](mailitem-votingoptions-property-outlook.md)**
    


If you attempt to use one of the computed properties listed below in a Jet filter (referencing the property by its explicit property name) for  **Table.Restrict**, you will get the error,  **IDS_ERR_ES_INVALIDRESTRICTION**: 


-  **AutoResolvedWinner**
    
-  **Body**
    
-  **BodyFormat**
    
-  **Class**
    
-  **ContactNames**
    
-  **Companies**
    
-  **[CompanyLastFirstNoSpace](contactitem-companylastfirstnospace-property-outlook.md)**
    
-  **[CompanyLastFirstSpaceOnly](contactitem-companylastfirstspaceonly-property-outlook.md)**
    
-  **ContactNames**
    
-  **[Contents](outlookbarpane-contents-property-outlook.md)**
    
-  **ConversationIndex**
    
-  **[DLName](distlistitem-dlname-property-outlook.md)**
    
-  **DownloadState**
    
-  **[Email1EntryID](contactitem-email1entryid-property-outlook.md)**
    
-  **[Email2EntryID](contactitem-email2entryid-property-outlook.md)**
    
-  **[Email3EntryID](contactitem-email3entryid-property-outlook.md)**
    
-  **EntryID**
    
-  **HtmlBody**
    
-  **InternetCodePage**
    
-  **IsConflict**
    
-  **IsMarkedAsTask**
    
-  **[LastFirstAndSuffix](contactitem-lastfirstandsuffix-property-outlook.md)**
    
-  **[LastFirstNoSpace](contactitem-lastfirstnospace-property-outlook.md)**
    
-  **[LastFirstNoSpaceAndSuffix](contactitem-lastfirstnospaceandsuffix-property-outlook.md)**
    
-  **[LastFirstNoSpaceCompany](contactitem-lastfirstnospacecompany-property-outlook.md)**
    
-  **[LastFirstSpaceOnly](contactitem-lastfirstspaceonly-property-outlook.md)**
    
-  **[LastFirstSpaceOnlyCompany](contactitem-lastfirstspaceonlycompany-property-outlook.md)**
    
-  **MeetingWorkspaceURL**
    
-  **MemberCount**
    
-  **[NetMeetingAlias](contactitem-netmeetingalias-property-outlook.md)**
    
-  **NetMeetingServer**
    
-  **[Permission](mailitem-permission-property-outlook.md)**
    
-  **[PermissionService](mailitem-permissionservice-property-outlook.md)**
    
-  **[RecurrenceState](appointmentitem-recurrencestate-property-outlook.md)**
    
-  **[ReceivedByEntryID](mailitem-receivedbyentryid-property-outlook.md)**
    
-  **[ReceivedOnBehalfOfEntryID](mailitem-receivedonbehalfofentryid-property-outlook.md)**
    
-  **ReplyRecipients**
    
-  **[ResponseState](taskitem-responsestate-property-outlook.md)**
    
-  **Saved**
    
-  **Sent**
    
-  **Submitted**
    
-  **TaskSubject**
    
-  **[VotingOptions](mailitem-votingoptions-property-outlook.md)**
    

 **Note**  For a computed property such as  **TaskSubject** or **IsMarkedAsTask**, you cannot add it to a  **Table** using **Columns.Add** or filter it using **Table.Restrict**, if you reference the property with the explicit property name. However, you can add or filter on the property if you reference it by namespace, as in the following code sample in Visual Basic for Applications: 



```vb
Sub TableForIsMarkedAsTask() 
    Dim oT As Outlook.Table 
    Dim oRow As Outlook.Row 
    Dim filter As String 
    '0x0E2B0003 represents IsMarkedAsTask 
    filter = "@SQL=" &; Chr(34) _ 
    &; "http://schemas.microsoft.com/mapi/proptag/0x0E2B0003" &; Chr(34) &; " = 1" 
    'Table only contains rows for items where IsMarkedAsTask is True 
    Set oT = Application.Session.GetDefaultFolder(olFolderInbox).GetTable(filter) 
    oT.Columns.Add ("TaskStartDate") 
    oT.Columns.Add ("TaskDueDate") 
    oT.Columns.Add ("TaskCompletedDate") 
    'Use GUID/ID to represent TaskSubject 
    oT.Columns.Add ( _ 
        "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85A4001E") 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        Debug.Print oRow( _ 
        "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85A4001E"), _ 
        oRow("TaskStartDate"), oRow("TaskDueDate"), oRow("TaskCompletedDate") 
    Loop 
End Sub
```


