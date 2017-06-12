---
title: Modify an Attachment of an Outlook Email Message
ms.prod: outlook
ms.assetid: f5dac09a-272b-49d6-bf1e-82c3981260ed
ms.date: 06/08/2017
---

# Modify an Attachment of an Outlook Email Message
This topic describes how you can programmatically modify a Microsoft Outlook email attachment without changing the original file.

 **Provided by:** Ken Getz, [MCW Technologies, LLC](http://www.mcwtech.com/)

Sending an email message with one or more attachments is easy, both in the Outlook interface and programmatically. In some scenarios, however, you may want to be able to modify the attachment after it has been attached to the mail item, without changing the original file in the file system. In other words, you may need a programmatic way to access the contents of the attachment in memory. 

For example, imagine that your application requires converting the text in all attachments that have a .txt extension to uppercase. In a managed Outlook add-in, you can easily handle the  **ItemSend** event. In that event, perform the work before sending the mail item. The difficult part of the scenario is in retrieving the contents of the attachments to modify the contents of each text file.

The sample code in this topic demonstrates how to solve this particular problem, using the  **GetProperty(String)** and **SetProperty(String, Object)** methods of the **Attachment** interface. In each case, you supply a value that contains the MAPI property [PidTagAttachDataBinary](http://msdn.microsoft.com/library/3b0a8b28-863e-4b96-a4c0-fdb8f40555b9%28Office.15%29.aspx) to obtain (and then set) the contents of the attachment.

 **Note**  The namespace representation of the  **PidTagAttachDataBinary** property is http://schemas.microsoft.com/mapi/proptag/0x37010102. For more information about using the **PropertyAccessor** object on properties that are referenced by namespace, see [Referencing Properties by Namespace](referencing-properties-by-namespace.md).

The sample code handles the  **ItemSend** event of a mail item. In the custom event handler, for any attachment that has a .txt extension, the code calls the `ConvertAttachmentToUpperCase` method. `ConvertAttachmentToUpperCase` takes an **Attachment** object and a **MailItem** object as input arguments, retrieves a byte array that is filled with the contents of the attachment, converts the byte array to a string, converts the string to uppercase, and then sets the contents of the attachment to the converted string as a byte array.

The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. 

You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code shows how to programmatically modify an Outlook email attachment without changing the original file. To demonstrate this functionality, in Visual Studio, create a new managed Outlook add-in named  `ModifyAttachmentAddIn`. Replace the code in ThisAddIn.cs or ThisAddIn.vb with the following code.

**Note** To  access the attachment data, the mail item must be saved using the [MailItem.Save](https://github.com/OfficeDev/VBA-content/edit/master/VBA/Outlook-VBA/articles/7d7b5f22-4749-e908-41a7-12a4c730c695.md) method.



```C#
using Outlook = Microsoft.Office.Interop.Outlook;
 
namespace ModifyAttachmentAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }
 
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
 
 
        void Application_ItemSend(object Item, ref bool Cancel)
        {
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
 
            if (mailItem != null)
            {
                var attachments = mailItem.Attachments;
                // If the attachment a text file, convert its text to all uppercase.
                foreach (Outlook.Attachment attachment in attachments)
                {

                    ConvertAttachmentToUpperCase(attachment, mailItem);
                }
            }
        }
 
        private void ConvertAttachmentToUpperCase(Outlook.Attachment attachment, Outlook.MailItem mailItem)
        {
            const string PR_ATTACH_DATA_BIN =
                "http://schemas.microsoft.com/mapi/proptag/0x37010102";
 
            // Confirm that the attachment is a text file.
            if (System.IO.Path.GetExtension(attachment.FileName) == ".txt")
            {
                // There are other heuristics you could use to determine whether the 
                // the attachment is a text file. For now, keep it simple: Only
                // run this code for *.txt.
 
                // Retrieve the attachment as an array of bytes.
                var attachmentData =
                    attachment.PropertyAccessor.GetProperty(
                    PR_ATTACH_DATA_BIN);
 
                // Convert the byte array into a Unicode string.
                string data = System.Text.Encoding.Unicode.GetString(attachmentData);
                // Convert to upper case.
                data = data.ToUpper();
                // Convert the data back to an array of bytes.
                attachmentData = System.Text.Encoding.Unicode.GetBytes(data);
 
                //Set PR_ATTACH_DATA_BIN to attachmentData.
                attachment.PropertyAccessor.SetProperty(PR_ATTACH_DATA_BIN,
                    attachmentData);
            }
        }
 
        #region VSTO generated code
 
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
 
        #endregion
    }
}
```




```VB.net
Public Class ThisAddIn
 
 
    Private Sub ThisAddIn_Startup() Handles Me.Startup
 
    End Sub
 
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
 
    End Sub
 
    Private Sub Application_ItemSend(ByVal Item As Object, _
        ByRef Cancel As Boolean) Handles Application.ItemSend
 
        Dim mailItem As Outlook.MailItem = TryCast(Item, Outlook.MailItem)
 
        If mailItem IsNot Nothing Then
            Dim attachments = mailItem.Attachments
            For Each attachment As Outlook.Attachment In attachments
                ' If the attachment is a text file, convert to uppercase.
                ConvertAttachmentToUpperCase(attachment, mailItem)
            Next attachment
        End If
    End Sub
 
    Private Sub ConvertAttachmentToUpperCase(ByVal attachment As Outlook.Attachment, _
        ByVal mailItem As Outlook.MailItem)
        Const PR_ATTACH_DATA_BIN As String = "http://schemas.microsoft.com/mapi/proptag/0x37010102"
 
        ' Confirm that the attachment is a text file.
        If System.IO.Path.GetExtension(attachment.FileName) = ".txt" Then
 
            ' There are other heuristics you could use to determine whether the 
            ' the attachment is a text file. For now, keep it simple: Only
            ' run this code for *.txt.
 
            ' Retrieve the attachment as an array of bytes.
            Dim attachmentData = attachment.PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN)
 
            ' Convert the byte array into a Unicode string.
            Dim data As String = System.Text.Encoding.Unicode.GetString(attachmentData)
            ' Convert to upper case.
            data = data.ToUpper()
            ' Convert the data back to an array of bytes.
            attachmentData = System.Text.Encoding.Unicode.GetBytes(data)
 
            'Set PR_ATTACH_DATA_BIN to attachmentData.
            attachment.PropertyAccessor.SetProperty(PR_ATTACH_DATA_BIN, attachmentData)
         End If
    End Sub
 
End Class
```


## See also


#### Concepts


 [Attach a File to a Mail Item](attach-a-file-to-a-mail-item.md)<br>
 [Attach an Outlook Contact Item to an Email Message](attach-an-outlook-contact-item-to-an-email-message.md)<br>
 [Limit the Size of an Attachment to an Outlook Email Message](limit-the-size-of-an-attachment-to-an-outlook-email-message.md)<br>

