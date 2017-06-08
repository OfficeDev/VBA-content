---
title: Save Auto-Archive Properties of a Folder in Solution Storage
ms.prod: outlook
ms.assetid: fbcdbbdf-3320-85f3-2dae-200fddd67285
ms.date: 06/08/2017
---


# Save Auto-Archive Properties of a Folder in Solution Storage

This topic shows a solution that saves its private data in a few MAPI auto-archive properties. The solution stores these properties in a  **[StorageItem](storageitem-object-outlook.md)** object of the folder to which the auto-archive properties apply. **StorageItem** objects are stored as hidden data in the associated portion of a folder, and because solutions can optionally encrypt their data, they offer the privacy required of solution data. Because the MAPI auto-archive properties are not exposed as explicit built-in properties in the Outlook object model, the solution uses the **[PropertyAccessor](propertyaccessor-object-outlook.md)** on the **StorageItem** object to set these properties.

The following illustrates the procedure:

1. The  `ChangeAgingProperties` function accepts the following as input parameters:
    
      -  `oFolder` is the **[Folder](folder-object-outlook.md)** object to which the aging properties apply and where their values are stored.
    
  -  `AgeFolder` indicates whether to archive or delete items in the folder as specified.
    
  -  `DeleteItems` indicates whether to delete, instead of archive, items that are older than the aging period.
    
  -  `FileName` indicates a specific file for archiving aged items. If this is an empty string, the default archive file, archive.pst, will be used.
    
  -  `Granularity` indicates the unit of time for aging, whether archiving is to be calculated in units of months, weeks, or days.
    
  -  `Period` indicates the amount of time in the given granularity. Together, the `Granularity` and `Period` values indicate an aging period. Items in the given folder older than this aging period are to be archived or deleted as specified. For example, a `Granularity` of 2 and `Period` of 14 specifies an aging period of 14 days, when items in the given folder older than 14 days should be either archived or deleted as specified.
    
  -  `Default` indicates which settings should be set to the default. The possible values are 0, 1, and 3:
    
      - 0 indicates nothing assumes a default value.
    
  - 1 indicates that only the file location assumes a default value. This is the same as checking  **Archive this folder using these settings** and **Move old items to default archive folder** in the **AutoArchive** tab of the **Properties** dialog box for the folder.
    
  - 3 indicates all settings assume a default value. This is the same as checking  **Archive items in this folder using default settings** in the **AutoArchive** tab of the **Properties** dialog box for the folder.
    
2. The validity of the parameters is checked.
    
3. If the parameters are valid,  ** [Folder.GetStorage](folder-getstorage-method-outlook.md)** is used to create or get an existing **StorageItem** object with the message class, **IPC.MS.Outlook.AgingProperties**. 
    
4.  **PropertyAccessor** is then used to set the auto-archive properties on the **StorageItem**,  ** [StorageItem.Save](storageitem-save-method-outlook.md)** is used to save changes to the **StorageItem**.
    
5. The  `TestAgingProps` procedure sets the auto-archive settings for the aging properties of the current folder so that items older than six months are moved to the default archive file.
    


## Remarks


1. Place the code in the built-in  **ThisOutlookSession** module.
    
2. Run the  `TestAgingProps` procedure to set aging properties on the current folder in the active explorer.
    

 **Note**  Whether it is implemented as a VBA macro or a COM add-in, the solution is a trusted caller and can therefore access the  **PropertyAccessor**. To improve this example, wrap the following VBA code in a .NET class for better error trapping and enumeration for  **Granularity**.


```vb
Function ChangeAgingProperties(oFolder As Outlook.Folder, _ 
 AgeFolder As Boolean, DeleteItems As Boolean, _ 
 FileName As String, Granularity As Integer, _ 
 Period As Integer, Default As Integer) As Boolean 
 
 '6 MAPI properties for aging items in a folder 
 Const PR_AGING_AGE_FOLDER = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x6857000B" 
 Const PR_AGING_DELETE_ITEMS = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x6855000B" 
 Const PR_AGING_FILE_NAME_AFTER9 = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x6859001E" 
 Const PR_AGING_GRANULARITY = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x36EE0003" 
 Const PR_AGING_PERIOD = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x36EC0003" 
 Const PR_AGING_DEFAULT = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x685E0003" 
 
 Dim oStorage As StorageItem 
 Dim oPA As PropertyAccessor 
 
 ' Valid Period: 
 ' 1-999 
 ' 
 ' Valid Granularity: 
 ' 0=Months, 1=Weeks, 2=Days 
 ' 
 ' Valid Default: 
 ' 0=All settings do not use a default setting 
 ' 1=Only the file location is defaulted 
 ' "Archive this folder using these settings" and 
 ' "Move old items to default archive folder" are checked 
 ' 3=All settings are defaulted 
 ' "Archive items in this folder using default settings" is checked 
 
 If (oFolder Is Nothing) Or _ 
 (Granularity < 0 Or Granularity > 2) Or _ 
 (Period < 1 Or Period > 999) Or _ 
 (Default < 0 Or Default = 2 Or Default > 3) _ 
 Then 
 ChangeAgingProperties = False 
 End If 
 
 On Error GoTo Aging_ErrTrap 
 
 'Create or get solution storage in given folder by message class 
 Set oStorage = oFolder.GetStorage( _ 
 "IPC.MS.Outlook.AgingProperties", olIdentifyByMessageClass) 
 Set oPA = oStorage.PropertyAccessor 
 
 If Not (AgeFolder) Then 
 oPA.SetProperty PR_AGING_AGE_FOLDER, False 
 Else 
 'Set the 6 aging properties in the solution storage 
 oPA.SetProperty PR_AGING_AGE_FOLDER, True 
 oPA.SetProperty PR_AGING_GRANULARITY, Granularity 
 oPA.SetProperty PR_AGING_DELETE_ITEMS, DeleteItems 
 oPA.SetProperty PR_AGING_PERIOD, Period 
 If FileName <> "" Then 
 oPA.SetProperty PR_AGING_FILE_NAME_AFTER9, FileName 
 End If 
 oPA.SetProperty (PR_AGING_DEFAULT), Default 
 End If 
 'Save changes as hidden messages to the associated portion of the folder 
 oStorage.Save 
 ChangeAgingProperties = True 
 Exit Function 
 
Aging_ErrTrap: 
 Debug.Print Err.Number, Err.Description 
 ChangeAgingProperties = False 
End Function 
 
Sub TestAgingProps() 
 Dim oFolder As Outlook.Folder 
 Set oFolder = Application.ActiveExplorer.CurrentFolder 
 If ChangeAgingProperties(oFolder, True, False, "", 0, 6, 1) Then 
 Debug.Print "ChangeAgingProperties OK" 
 Else 
 Debug.Print "ChangeAgingProperties Failed" 
 End If 
End Sub
```


