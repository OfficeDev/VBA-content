---
title: List the Name and Office Location of Each Manager Belonging to an Exchange Distribution List
ms.prod: outlook
ms.assetid: abc26854-62db-be7f-4025-46acbcb42541
ms.date: 06/08/2017
---


# List the Name and Office Location of Each Manager Belonging to an Exchange Distribution List

This topic describes how to allow a user to select an Exchange distribution list and display the name and office location of each member who is a manager belonging to that distribution list. The major steps of this procedure are as follows:


1. The code sample below displays a  **Select Distribition List** dialog box for the user to select a distribution list.
    
    It uses the  **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object to display the dialog box and obtain user selection. The sample then obtains the user selection through the **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)** property.
    
2. For each member in the selected distribution list:
    
      1. If the member is a manager, then the code sample displays the name and office number of the manager. 
    
    Each member in the distribution list is an  **[AddressEntry](addressentry-object-outlook.md)** object. By checking if the **[AddressEntry.AddressEntryUserType](addressentry-addressentryusertype-property-outlook.md)** is either **olExchangeUserAddressEntry** or **olExchangeRemoteUserAddressEntry**, the sample then assigns the  **AddressEntry** object to an **[ExchangeUser](exchangeuser-object-outlook.md)** object, and uses `ExchangeUser.GetDirectReports.Count >0` as a criterion to determine if the user is a manager. It then displays the **[Name](exchangeuser-name-property-outlook.md)** and **[OfficeLocation](exchangeuser-officelocation-property-outlook.md)** properties of the **ExchangeUser** object.
    
  2. If the member is a distribution list, the code sample calls the subroutine  `EnumerateDLManagers`. For each member in that distribution list, if the member is a manager, the code sample then displays the name and office number of the manager.
    

Copy the following Visual Basic for Applications code sample to the Visual Basic Editor, and run  `ShowManagersOfGroups`. Note that this code sample only applies to a distribution list that has only Exchange users as members, or that has Exchange distribution lists as members but all members of the latter will have to be Exchange users. Further customization of the code will be necessary if there is more nesting of distribution lists as members. 




```vb
Sub ShowManagersOfGroups() 
    Dim oRecip As Outlook.Recipient 
    Dim oSND As Outlook.SelectNamesDialog 
    Dim oAE As Outlook.AddressEntry 
    Dim oAEs As Outlook.AddressEntries 
    Dim oEU As Outlook.ExchangeUser 
    Dim oDL As Outlook.ExchangeDistributionList 
    Dim oLists As Outlook.AddressLists 
    Dim oList As Outlook.AddressList 
    Set oLists = Application.Session.AddressLists 
    For Each oList In oLists 
        If oList.Name = "All Groups" Then 
            Exit For 
        End If 
    Next 
    Set oSND = Application.Session.GetSelectNamesDialog 
    With oSND 
        .NumberOfRecipientSelectors = olShowTo 
        .InitialAddressList = oList 
        .Caption = "Select Distribution List" 
        .ToLabel = "D/L" 
        .ShowOnlyInitialAddressList = True 
        .AllowMultipleSelection = False 
        .Display 
    End With 
    For Each oRecip In oSND.Recipients 
        If oRecip.AddressEntry.AddressEntryUserType = _ 
            olExchangeDistributionListAddressEntry Then 
            Set oDL = oRecip.AddressEntry.GetExchangeDistributionList 
            Set oAEs = oDL.GetExchangeDistributionListMembers 
            For Each oAE In oAEs 
                If oAE.AddressEntryUserType = olExchangeUserAddressEntry _ 
                    Or oAE.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then 
                    Set oEU = oAE.GetExchangeUser 
                    If oEU.GetDirectReports.Count Then 
                        Debug.Print oEU.Name, oEU.OfficeLocation 
                    End If 
                ElseIf oAE.AddressEntryUserType = _ 
                    olExchangeDistributionListAddressEntry Then 
                    EnumerateDLManagers oAE 
                End If 
            Next 
        End If 
    Next 
End Sub 
 
Sub EnumerateDLManagers(oAddress As AddressEntry) 
    Dim oAE As Outlook.AddressEntry 
    Dim oAEs As Outlook.AddressEntries 
    Dim oEU As Outlook.ExchangeUser 
    Dim oDL As Outlook.ExchangeDistributionList 
     
    Set oDL = oAddress.GetExchangeDistributionList 
    Set oAEs = oDL.GetExchangeDistributionListMembers 
    For Each oAE In oAEs 
        If oAE.AddressEntryUserType = olExchangeUserAddressEntry _ 
            Or oAE.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then 
            Set oEU = oAE.GetExchangeUser 
            If oEU.GetDirectReports.Count Then 
                Debug.Print oEU.Name, oEU.OfficeLocation 
            End If 
        End If 
    Next 
End Sub
```


