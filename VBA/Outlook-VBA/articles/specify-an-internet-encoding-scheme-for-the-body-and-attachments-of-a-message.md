---
title: Specify an Internet Encoding Scheme for the Body and Attachments of a Message
ms.prod: outlook
ms.assetid: e6207bf2-238d-2b7a-cd80-5783e49c05ec
ms.date: 06/08/2017
---


# Specify an Internet Encoding Scheme for the Body and Attachments of a Message

This topic shows how to use the MAPI property,  [PidTagInternetMailOverrideFormat](http://msdn.microsoft.com/library/guid_2fc91e13-703c-3ec9-9066-ffee7144306c.aspx), and the Microsoft Outlook object model to specify an Internet encoding scheme for the Exchange Internet Mail Service (IMS) to encode the body and attachments of a mail item.

The following code sample in Visual C# shows how to reference  **PidTagInternetMailOverrideFormat** with its MAPI proptag namespace and use the **[PropertyAccessor](http://msdn.microsoft.com/library/guid_2fc91e13-703c-3ec9-9066-ffee7144306c.aspx)** object of the Outlook object model to specify MIME as the Internet encoding scheme for a message. **PidTagInternetMailOverrideFormat** is referenced as:



```
http://schemas.microsoft.com/mapi/proptag/0x59020003
```

where  `0x59020003` is the proptag of **PidTagInternetMailOverrideFormat**.



```C#
private void SendMail_Click() 
{ 
    Outlook.NameSpace objSession; 
    Outlook.MailItem objMailItem; 
    Outlook.Recipient objRecipient; 
    Outlook.PropertyAccessor oPA; 
 
    string Recipient, MsgSubject, ImageFile, TextFile, FileLocation, PropName; 
    int EncodingFlag; 
     
 
    //Modify the following to appropriate values. 
    Recipient = "someone@example.com"; 
    EncodingFlag = 1; //Use MIME encoding 
    MsgSubject = "Test Encoding"; 
    ImageFile = "garden.jpg"; 
    TextFile = "mytext.txt"; 
    FileLocation = "c:\\"; 
 
    objSession = Application.GetNamespace("MAPI"); 
    objSession.Logon(null, null, true, null); 
 
    objMailItem = Application.CreateItem( 
                Outlook.OlItemType.olMailItem) as Outlook.MailItem; 
    objMailItem.Subject = MsgSubject; 
    objMailItem.Body = "body"; 
    objMailItem.Attachments.Add(FileLocation + TextFile,  
        Outlook.OlAttachmentType.olByValue, 1, TextFile); 
    objMailItem.Attachments.Add(FileLocation + ImageFile, 
        Outlook.OlAttachmentType.olByValue, 1, ImageFile); 
 
    objRecipient = objMailItem.Recipients.Add(Recipient); 
    objRecipient.Resolve(); 
 
    PropName = "http://schemas.microsoft.com/mapi/proptag/0x59020003"; 
    oPA = objMailItem.PropertyAccessor; 
    oPA.SetProperty(PropName, EncodingFlag); 
 
    objMailItem.Send(); 
 
    objSession.Logoff(); 
 
}

```


