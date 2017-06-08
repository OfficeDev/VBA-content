---
title: Make Separate Compose and Read Versions for a Form
ms.prod: outlook
ms.assetid: 6c533327-ce16-169c-6c6a-dd6cecb0e3fb
ms.date: 06/08/2017
---


# Make Separate Compose and Read Versions for a Form

You can create and edit separate compose and read versions for each page of a form. The compose version is the one that is displayed when you create an item. The read version is the one that Outlook displays after the item has been saved—for example, what Outlook displays in an inspector window when you read your email.


 **Note**  Outlook only supports separate compose and read versions on some item types; in particular, items that you can send, such as mail and post. Other item types always display the compose form.


## Forms customized with form regions

When you use form regions to create separate compose and read versions for a form, you must design two separate form regions, one for each version, and then register them independently on the same message class. Properties that you set in the XML manifest indicate which form is for the read-only view and which should be used for compose.

For more information, see  [How to: Create Separate Read and Compose Pages Using Form Regions in a Custom Form](create-separate-read-and-compose-pages-using-form-regions-in-a-custom-form.md).


## Forms customized with form pages


1. In Forms Designer, on the  **Developer** tab, in the **Form** group, do one of the following:
    
      - To create or edit the compose page, click  **Page**, then click  **Edit Compose Page**.
    
  - To create or edit the read page, click  **Page**, then click  **Edit Read Page**.
    
2. Customize pages on the form.
    
    For more information, see  [How to: Customize Pages on a Form](customize-pages-on-a-form.md).
    

 **Note**  You can quickly switch between the two versions by clicking  **Edit Compose Page** and **Edit Read Page** in the **Design** group. If these commands are not available, you have set the compose and read versions to be the same. To have separate compose and read versions, click **Separate Read Layout** in the **Design** group.


