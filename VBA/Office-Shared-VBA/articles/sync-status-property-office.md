---
title: Sync.Status Property (Office)
keywords: vbaof11.chm277001
f1_keywords:
- vbaof11.chm277001
ms.prod: office
api_name:
- Office.Sync.Status
ms.assetid: fdddff38-268b-835a-7c8d-db76d862e392
ms.date: 06/08/2017
---


# Sync.Status Property (Office)

Gets the status of the synchronization of the local copy of the active document with the server copy. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Status**

 _expression_ Required. A variable that represents a **[Sync](sync-object-office.md)** object.


## Remarks

Use the  **Status** property to determine whether the local copy of the active document is synchronized with the shared server copy. Use the **GetUpdate** method to refresh the status. Use the following methods and properties when appropriate to respond to various status conditions:




-  ** msoSyncStatusConflict** - **True** when both the local and the server copies have changes. Use the **ResolveConflict** method to resolve the differences.
    
-  **msoSyncStatusError** - Check the **ErrorType** property.
    
-  ** msoSyncStatusLocalChanges** - **True** when only the local copy has changes. Use the **PutUpdate** method to save local changes to the server copy.
    
-  ** msoSyncStatusNewerAvailable** - **True** when only the server copy has changes. Close and re-open the document to work with the latest copy from the server.
    
-  ** msoSyncStatusSuspended** - Use the **Unsuspend** method to resume synchronization.
    


The  **Status** property returns a single constant from the list in the following order of precedence:


1.  **msoSyncStatusNoSharedWorkspace**
    
2.  **msoSyncStatusError**
    
3.  **msoSyncStatusSuspended**
    
4.  **msoSyncStatusConflict**
    
5.  **msoSyncStatusNewerAvailable**
    
6.  **msoSyncStatusLocalChanges**
    
7.  **msoSyncStatusLatest**
    



## Example

The following example examines the  **Status** property and takes an appropriate action to synchronize the local and server copies of the document if necessary.


```
    Dim objSync As Office.Sync 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    If objSync.Status > msoSyncStatusNoSharedWorkspace Then 
        Select Case objSync.Status 
            Case msoSyncStatusConflict 
                objSync.ResolveConflict msoSyncConflictMerge 
                ActiveDocument.Save 
                objSync.ResolveConflict msoSyncConflictClientWins 
                strStatus = "Conflict resolved by merging changes." 
            Case msoSyncStatusError 
                strStatus = "Last error type: " &amp; objSync.ErrorType 
            Case msoSyncStatusLatest 
                strStatus = "Document copies already in sync." 
            Case msoSyncStatusLocalChanges 
                objSync.PutUpdate 
                strStatus = "Local changes saved to server." 
            Case msoSyncStatusNewerAvailable 
                strStatus = "Newer copy available on the server." 
            Case msoSyncStatusSuspended 
                objSync.Unsuspend 
                strStatus = "Synchronization resumed." 
        End Select 
    Else 
        strStatus = "Not a shared workspace document." 
    End If 
    MsgBox strStatus, vbInformation + vbOKOnly, "Sync Information" 
    Set objSync = Nothing 

```


## See also


#### Concepts


[Sync Object](sync-object-office.md)
#### Other resources


[Sync Object Members](sync-members-office.md)

