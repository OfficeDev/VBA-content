---
title: Setting Default Properties for a Form
ms.prod: outlook
ms.assetid: dd3dd5c6-bc23-70d5-ae6c-b8a2bb4e9a66
ms.date: 06/08/2017
---


# Setting Default Properties for a Form

To set the default properties for a custom form that has form pages, use the  **Properties** tab in the Forms Designer.


 **Note**  For more information about how to set the default properties in a Microsoft Outlook form that has form regions as opposed to form pages, see  [How to: Create a Form Region](create-a-form-region.md).


The following are the default properties for custom forms that have form pages:


-  **Category** You can specify a category for your form to help organize the forms in the **Choose Form** dialog box when you select a form.
    
-  **Sub-Category** You can further refine the category by specifying a sub-category.
    
-  **Always use Microsoft Word as the e-mail editor** Since Microsoft Office Outlook 2007, Outlook uses Microsoft Word as the e-mail editor. However, using Word as the editor is optional in earlier versions. If you create forms for users who use earlier versions of Outlook, you can use this option to specify Word as the editor for the message portion (or control) of your form. Users then have all the formatting options that are available with Word, such as the spell checker and thesaurus. For these options to be available, the recipients of your form must have Word installed.
    
     **Note**  This feature has not been changed from earlier versions of Microsoft Office. The code behind this option uses an older architecture for using Word as the e-mail editor and does not provide the same user experience with Word that clicking  **Options** on the Microsoft Office Backstage view does. Solutions that you create for earlier versions using Word as the e-mail editor might not work, or might not work properly, in Office Outlook 2007 or a later version of Outlook.
-  **Template** You can specify the Word template to use to format the text in the message control of the form.
    
     **Note**  Although Word is now the e-mail editor, this setting, as applicable to Outlook forms, is enabled only if you select the  **Always use Microsoft Word as the e-mail editor** check box. If this check box is cleared, you cannot set templates.
-  **Contact** When you click **Contact**, you can access the Address Book, where you can select the people who maintain, upgrade, or distribute information about this form. The contact information that you provide appears in the  **Forms Manager** dialog box and the form **Properties** page.
    
-  **Description** You can type a description for your form. For example, you could include instructions for how to use the form or explain the purpose of the form. Outlook displays the description in the **About** dialog box on the **Help** menu of the form and in the **Properties** dialog box for the form.
    
-  **Version** You can set a version number for this form. This is a free-form text field and does not affect Outlook behavior in any way.
    
-  **Form Number** You can set a unique form number that identifies the form. This is a free-form text field and does not affect Outlook behavior in any way.
    
-  **Change Large Icon** Click this button to open the **File Open** dialog box, where you can select a different large icon for your form. Large icons appear in the form **Properties** dialog box.
    
-  **Change Small Icon** Click this button to open the **File Open** dialog box, where you can select a different small icon for your form. Small icons appear in the Outlook folder to represent an item of the type the form creates.
    
-  **Send form definition with item** Instructs Outlook to include the form definition when you send the form. (Note that the form is much larger when it includes the form definition.) When you select this option, Outlook creates a self-contained form that the recipients can use to view the form, even if they do not have access to the same forms library as the sender.
    
     **Note**  Because of improvements to security since Office Outlook 2007, this option is not recommended nor is it necessary in most cases. In general, publishing the form is all that is needed.

    Outlook does not run Microsoft Visual Basic Scripting Edition (VBScript) code if the form definition is included with the item. In most cases, it is better to publish a form instead of including the form definition with the item. If you do send the form with the item, you can re-enable the VBScript code if you use the Outlook custom security settings. In that case, if you send a form with this box checked, the recipients see a  **Warning** dialog box, and they have the option to disable the macros because the form is not published. Harmful macros could delete or copy their files, or send e-mail messages from their mailbox to another user.If network or file transfer time is an issue, and you cannot publish the form for some reason, an alternative to sending the form definition is to save the form and send it as an attachment to another form. Recipients can take the attached form and publish it in their own forms library.
    
-  **Use form only for responses** Hides a form when it is published to a forms library. This option is useful in situations when you create a form to use solely for replies. In another form, you can specify that your reply form will be used instead of the default reply form.To use your form only for responses, select the **Use form only for responses** check box, and then [publish your form](publish-a-form.md). Open a second form in design mode. On the  **Actions** page of the second form, you can specify your published form in the **Reply** or **Reply to All** action. To use your form as the default reply form, double-click the **Reply** action in the second form. You can select the name of your published reply form in the **Form name** field of the **Form Action Properties** dialog box.
    

