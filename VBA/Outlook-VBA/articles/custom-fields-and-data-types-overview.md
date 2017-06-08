---
title: Custom Fields and Data Types Overview
ms.prod: outlook
ms.assetid: a85a7bc2-2b85-1782-04a3-0104e0df32aa
ms.date: 06/08/2017
---


# Custom Fields and Data Types Overview

You can use custom fields in several ways in Microsoft Outlook, including the following:


- You can create new data-entry fields for a view or a form so that users can add their own custom information. For example, you can create a  **Date/Time** field named "Last Talked To" in a contact folder and then add the field to a **Phone List** view.
    
- You can create new views by combining the standard fields in a single column. For example, you can create a column that combines the  **State** and **Country** fields in an address list to save space.
    
- You can create a formula field to show information in a new way. For example, you can create a field that displays "Large" when a message is more than 10,000 bytes and "Small" when a message is under the same limit.
    

Any field that you create is stored in the folder where you create it. To use a field in more than one folder, you must create the field in each folder.

You can create and view custom fields in table views and card views. You can create custom fields with the following data types in Outlook.


|**Data type**|**Use to represent**|
|:-----|:-----|
| **Combination**|Combinations of values of fields and text in a column (table) or row (card). You can specify whether to show each field or show only the first non-empty field. You can also combine text with fields without the use of quotation marks. |
| **Currency**|Numeric data as currency or mathematical calculations that involve money.|
| **Date/Time**|Date and time data.|
| **Duration**|Numeric data. You can enter a duration as minutes, hours, or days. Values are saved as minutes.|
| **Formula**|Calculations based on standard and custom fields. You can use any appropriate functions and operators to complete the formula. |
| **Integer**|Non-decimal numeric data.|
| **Keywords**|When filled in, this user-defined field is used to group and find related items similar to the way the  **Categories** field is used in Outlook. Text with multiple values to be separated by commas. Each value can be grouped individually in a view.|
| **Number**|Numeric data or mathematical calculations except those that involve money. (For money, use  **Currency** data type.)|
| **Percent**|Numeric data as a percentage.|
| **Text**|Text or combinations of text and numbers, such as addresses. Can be up to 255 characters long.|
| **Yes/No**|Data that contains only one of two values, such as Yes/No, True/False, On/Off. |
Each of the data types except  **Combination**,  **Formula**, and  **Keywords** has a series of standard formats that you can use to show the values of the fields.

