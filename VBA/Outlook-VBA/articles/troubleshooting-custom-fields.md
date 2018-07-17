---
title: Troubleshooting Custom Fields
ms.prod: outlook
ms.assetid: 0867a1a9-595b-4dd0-becc-e1ba744a76c1
ms.date: 06/08/2017
---


# Troubleshooting Custom Fields

Use the following troubleshooting tips to help troubleshoot problems with custom fields.


## The custom field I created is not visible.

 **Forms customized with form pages**


- Custom fields are stored in the folder in which you create them. Open the folder in which you created the custom field to see if it is visible. You must recreate a custom field for each folder in which you want to use it.
    
- Custom fields are stored in the field set called  **User-defined** fields in the **Field Chooser** and in the **Show Fields** dialog box. To see all the custom fields in a folder, do the following:
    
      1. Open the folder.
    
  2. On the  **View** menu, select **Current View**, and then click  **Customize Current View**.
    
  3. In the  **Customize View** dialog box, click **Fields**.
    
  4. In the  **Select available fields from** box, click **User-defined fields in folder**.
    
 **Forms customized with form regions**


- When using a custom field on a form region, you must create the field by using an add-in in each folder where you store items that use the form. That way, you can add the field to the view and it is searchable. If you create the field on only the form region, it will not be accessible in the view and it might not pick up the appropriate default value when you create a new item that uses the form.For more information, see  [Extending a Form Region with an Add-in](extending-a-form-region-with-an-add-in.md).
    

## When I try to sort, group, or filter a formula field or combination field, I receive an error message.


- You cannot sort, group, or filter a formula field or a combination field in Microsoft Outlook.
    

