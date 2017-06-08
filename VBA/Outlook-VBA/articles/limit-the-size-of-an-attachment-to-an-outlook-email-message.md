---
title: Limit the Size of an Attachment to an Outlook Email Message
ms.prod: outlook
ms.assetid: 9a240e17-f715-482c-9a8b-c6be1144e15a
ms.date: 06/08/2017
---


# Limit the Size of an Attachment to an Outlook Email Message
This topic describes how you can create a managed add-in for Outlook that cancels sending email if the total attachment size is greater than a fixed limit.

 **Provided by:** Ken Getz, [MCW Technologies, LLC](http://www.mcwtech.com/)

A given email message can contain one or more file attachments, and you may want to limit the total attachment size in email messages that you send. The sample code in this topic demonstrates how you can handle the  **ItemSend** event in an Outlook add-in, and in the event handler, cancel the sending of the email message if the combined size of all the attachments is larger than a specific value (2 MB, in this example).

The Outlook  **ItemSend** event receives as its parameters a reference to the item being sent, and a Boolean variable that is passed by reference and that allows you to cancel the send operation. It is up to your own code in the event handler to determine whether you want to cancel the event; you do so by setting theCancel parameter to **True** if you do wish to cancel the event.

In this example, to determine whether the total attachment size is larger than a specific size, the code loops through each attachment in the item's  **Attachments** collection. For each item, the code retrieves the **Size** property, summing as it loops. If the sum ever exceeds the size of the `maxSize` constant, the code sets the `tooLarge` variable to **True**, and exits the loop. After the loop, if the  `tooLarge` variable is **True**, the code alerts the user and sets the Cancel parameter to the event handler (which was passed by reference) to **True**, causing Outlook to cancel the sending of the item.

The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. 

You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code shows how to cancel sending an email if the total attachment size is greater than the specified limit. To demonstrate this functionality, in Visual Studio, create a new managed Outlook add-in named  `LimitAttachmentSizeAddIn`. Replace the code in ThisAddIn.cs or ThisAddIn.vb with the example code shown here.



```C#
using Outlook = Microsoft.Office.Interop.Outlook;
 
namespace LimitAttachmentSizeAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
          Application.ItemSend +=new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }
 
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
 
        void Application_ItemSend(object Item, ref bool Cancel)
        {
            // Specify the maximum size for the attachments. For this example,
            // the maximum size is 2 MB.
            const int maxSize = 2 * 1024 * 1000;
            bool tooLarge = false;
 
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            if (mailItem != null)
            {
                var attachments = mailItem.Attachments;
                double totalSize = 0;
                foreach (Outlook.Attachment attachment in attachments)
                {
                    totalSize += attachment.Size;
                    if (totalSize > maxSize)
                    {
                        tooLarge = true;
                        break;
                    }
                }
            }
            if (tooLarge)
            {
                // If the sum of the attachment sizes is too large, alert the user
                // and cancel the send.
                System.Windows.Forms.MessageBox.Show(
                    "The total attachment size is too large. Sending canceled.", 
                    "Outlook Add-In");
                Cancel = true;
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
 
    Private Sub Application_ItemSend(ByVal Item As Object, ByRef Cancel As Boolean) Handles Application.ItemSend
        ' Specify the maximum size for the attachments. For this example,
        ' the maximum size is 2 MB.
        Const maxSize As Integer = 2 * 1024 * 1000
        Dim tooLarge As Boolean = False
 
        Dim mailItem As Outlook.MailItem = TryCast(Item, Outlook.MailItem)
        If mailItem IsNot Nothing Then
            Dim attachments = mailItem.Attachments
            Dim totalSize As Double = 0
 
            For Each attachment As Outlook.Attachment In attachments
                totalSize += attachment.Size
                If totalSize > maxSize Then
                    tooLarge = True
                    Exit For
                End If
            Next attachment
        End If
 
        If tooLarge Then
            ' If the sum of the attachment sizes is too large, alert the user
            ' and cancel the send.
            System.Windows.Forms.MessageBox.Show(
                "The total attachment size is too large. Sending canceled.",
                "Outlook Add-In")
            Cancel = True
        End If
    End Sub
End Class
```


## See also


#### Concepts


 [Attach a File to a Mail Item](attach-a-file-to-a-mail-item.md)<br>
 [Attach an Outlook Contact Item to an Email Message](attach-an-outlook-contact-item-to-an-email-message.md)<br>
 [Modify an Attachment of an Outlook Email Message](modify-an-attachment-of-an-outlook-email-message.md)

