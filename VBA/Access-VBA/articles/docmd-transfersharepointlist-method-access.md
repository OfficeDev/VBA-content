---
title: DoCmd.TransferSharePointList Method (Access)
keywords: vbaac10.chm5618
f1_keywords:
- vbaac10.chm5618
ms.prod: access
api_name:
- Access.DoCmd.TransferSharePointList
ms.assetid: 9cbd8de6-dc1a-47b0-c1f4-62959a66faf4
ms.date: 06/08/2017
---


# DoCmd.TransferSharePointList Method (Access)

You can use the  **TransferSharePointList** method to import or link data from a SharePoint Foundation site.


## Syntax

 _expression_. **TransferSharePointList**( ** _TransferType_**, ** _SiteAddress_**, ** _ListID_**, ** _ViewID_**, ** _TableName_**, ** _GetLookupDisplayValues_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TransferType_|Required|**AcSharePointListTransferType**|An  **[AcSharePointListTransferType](acsharepointlisttransfertype-enumeration-access.md)** constant that specifies the type of transfer to make.|
| _SiteAddress_|Required|**Variant**|The full path of the SharePoint site.|
| _ListID_|Required|**Variant**|The name or GUID of the list to be transferred.|
| _ViewID_|Optional|**Variant**|The GUID of the view for the list you want to use. Leave this argument blank to transfer all rows and columns in the list.|
| _TableName_|Optional|**Variant**|The name you want displayed for the table or linked table in Access.|
| _GetLookupDisplayValues_|Optional|**Variant**|Specifies whether to transfer display values for Lookup fields instead of the ID used to perform the lookup.|

## Remarks

This method has the same effect as clicking  **SharePoint List** in the **Import** group on the **External Data** tab. The arguments for the action correspond to the choices you make in the Get External Data Wizard.

If you specify a nonexistent list or view, no error occurs, and no data is transferred.

A GUID is a unique hexadecimal identifier for a list or a view. A GUID must be entered in the following format, where each "F" is a hexadecimal number (0 through 9 or A through F).




```
 
{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}
```

You can obtain the GUID for a list or view from the SharePoint site by using the following procedure:


1. Open the list in SharePoint Foundation.
    
2. If the view you want is not displayed, click the  **View** drop-down arrow and then select the view you want.
    
3. Click the  **View** drop-down arrow and then select **Modify this View**.The address in the browser's address bar contains the GUIDs for both the list and the view. The GUID for the list follows  **List=**, and the GUID for the view follows **View=**. However, in the address, each **{** (left brace) character is represented by the string **%7B**, each **-** (hyphen) character is represented by the string **%2D**, and each **}** (right brace) character is represented by the string **%7D**. For example: `http://MySite12/_layouts/ViewEdit.aspx?List=%7B2A82A404%2D5529%2D47DC%2DAE13%2DAC1D9BC0A84F%7D&;View=%7B357B4FE6%2D44CF%2D4275%2DB91F%2D46558301579B%7D`Before you can use the GUIDs from the address as arguments in this macro action, you must replace each  **%7B** string with the **{** character, replace each **%2D** string with the **-** character, and replace each **%7D** string with the **}** character. Do not include the **&;** (ampersand) character that follows the **%7D** string in the list GUID.
    

## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

