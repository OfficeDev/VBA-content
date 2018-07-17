---
title: Implementing the IRibbonExtensibility Interface
ms.prod: outlook
ms.assetid: ad798afe-b3a9-4d03-86b3-b1226d9b55c8
ms.date: 06/08/2017
---


# Implementing the IRibbonExtensibility Interface

A Microsoft Outlook add-in that customizes the ribbon, shortcut menus, new-item menus, or Microsoft Office Backstage view must implement the following interfaces:


-  **IDTExtensibility2**
    
-  **[IRibbonExtensibility](http://msdn.microsoft.com/library/b27a7576-b6f5-031e-e307-78ef5f8507e0%28Office.15%29.aspx)**
    

In Visual C# and Visual Basic add-ins, you must implement these interfaces in the same class.
When you implement  **Office.IRibbonExtensibility**, use the  **IRibbonExtensibility.GetCustomUI** method to return XML markup for your custom user interface to Outlook. The way that Outlook calls **GetCustomUI** and when it calls it is unique among Microsoft Office applications:

- Office calls  **GetCustomUI** during Outlook startup to load ribbon customizations for the explorers.
    
- Office calls  **GetCustomUI** to load inspector-specific ribbon customizations when the first instance of a given inspector type, such as an appointment or contact inspector, is displayed.
    
- Viewing an item in the Reading Pane does not cause  **GetCustomUI** to be called because the ribbon is not displayed in the Reading Pane.
    
The ribbon ID is a string that is passed from Office to  **GetCustomUI** and that specifies the UI customization to load. Add-in developers can use this string to determine the custom XML markup to return to Outlook. You can also use the ribbon ID to determine the type of Outlook item to display.
In some cases, such as a  **[MailItem](mailitem-object-outlook.md)** or **[PostItem](postitem-object-outlook.md)**, Outlook calls  **GetCustomUI** once when the first compose note is displayed (where `RibbonID = Microsoft.Outlook.Mail.Compose`) and another time when the first read note is displayed (where  `RibbonID = Microsoft.Outlook.Mail.Read`).
Outlook uses the following unique ribbon IDs.


| **Ribbon ID**| **Message Class**|
|:-----|:-----|
|Microsoft.OMS.MMS.Compose|IPM.Note.Mobile.MMS.*|
|Microsoft.OMS.MMS.Read|IPM.Note.Mobile.MMS.*|
|Microsoft.OMS.SMS.Compose|IPM.Note.Mobile.MMS.*|
|Microsoft.OMS.SMS.Read|IPM.Note.Mobile.MMS.*|
|Microsoft.Outlook.Appointment|IPM.Appointment.*|
|Microsoft.Outlook.Contact|IPM.Contact.*|
|Microsoft.Outlook.DistributionList|IPM.DistList.*|
|Microsoft.Outlook.Journal|IPM.Activity.*|
|Microsoft.Outlook.Mail.Compose|IPM.Note.*|
|Microsoft.Outlook.Mail.Read|IPM.Note.*|
|Microsoft.Outlook.MeetingRequest.Read|IPM.Schedule.Meeting.Request or IPM.Schedule.Meeting.Canceled|
|Microsoft.Outlook.MeetingRequest.Send|IPM.Schedule.Meeting.Request|
|Microsoft.Outlook.Post.Compose|IPM.Post.*|
|Microsoft.Outlook.Post.Read|IPM.Post.*|
|Microsoft.Outlook.Report|IPM.Report.*|
|Microsoft.Outlook.Resend|IPM.Resend.*|
|Microsoft.Outlook.Response.Compose|IPM.Schedule.Meeting.Resp.*|
|Microsoft.Outlook.Response.CounterPropose|IPM.Schedule.Meeting.Resp.*|
|Microsoft.Outlook.Response.Read|IPM.Schedule.Meeting.Resp.*|
|Microsoft.Outlook.RSS|IPM.Post.Rss|
|Microsoft.Outlook.Sharing.Compose|IPM.Sharing.*|
|Microsoft.Outlook.Sharing.Read|IPM.Sharing.*|
|Microsoft.Outlook.Task|IPM.Task.* and IPM.TaskRequest.*|
|Microsoft.Outlook.Explorer|Not applicable. This ribbon ID lets you return XML for explorer ribbons, shortcut menus, and Backstage view.|

 **Note**  Sticky notes do not implement the ribbon, so IPM.StickyNote is not listed in the table of ribbon IDs and message classes.

For all of the ribbon ID values except for Microsoft.Outlook.Explorer, notice that the corresponding message class is listed as **IPM.Type.** in the table. That means that either the first instance of the base message class (for example, IPM.Contact) or of a derived custom message class (IPM.Contact.ShoeStore) that appears in an inspector will cause Outlook to call  **GetCustomUI**. 

Because a base message class shares the same ribbon XML with the custom message classes that are derived from it, and because Outlook calls  **GetCustomUI** only one time per ribbon ID, you cannot specify a ribbon XML markup that is applied exclusively to a derived custom message class, and not to the base message class. However, if you want controls to appear in the ribbon on inspectors for one custom message class and not all other message classes with the same base message class, do the following:

1. In  **GetCustomUI**, return the XML markup for the ribbon ID for the custom message class (for example, IPM.Contact.ShoeStore) to the ribbon. All ribbons that are used by items with the same base message class (for example, IPM.Contact) will contain the added controls.
    
2.  In the ribbon XML, specify **[IRibbonControl.Context](http://msdn.microsoft.com/library/39f9d85a-00e9-9682-3957-51d9e72b4d83%28Office.15%29.aspx)** callbacks for each tab, group, and control that is specific to the custom message class. These callbacks will be used to display the controls for the custom message class and to hide the controls for the base message class and all other message classes with the same base message class.
    
3.  In each **getVisible** callback, cast the **IRibbonControl.Context** parameter that is passed to the callback to an Outlook **[Inspector](inspector-object-outlook.md)** object. Use the **MessageClass** property of **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** to determine whether to return **True** or **False** in the **getVisible** callback.
    
To customize the ribbon on all or multiple Outlook message classes, use the following recommendations:

- To customize the first built-in tab on all Outlook inspectors, you must supply separate ribbon XML for different ribbon IDs because built-in first tabs do not have the same name across all ribbon IDs.
    
- To customize the ribbon on multiple Outlook inspectors, you might have to supply separate ribbon XML for different ribbon IDs depending upon the tab name.
    
For more information, including examples, about customizing explorer and inspector ribbons, shortcut menus, and Backstage view, see  [Extending the User Interface in Outlook 2010](http://msdn.microsoft.com/library/00b504b0-e897-43b9-8615-44276166823f.aspx) on the MSDN Web site.

## See also


#### Concepts


 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)<br>
 [Overview of the IRibbonUI Object](overview-of-the-iribbonui-object.md)<br>
 [Detecting Errors](detecting-errors.md)<br>
 [Overview of the IRibbonControl Object](overview-of-the-iribboncontrol-object.md)

