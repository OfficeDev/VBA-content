---
title: EncryptionProvider Members (Office)
ms.prod: office
ms.assetid: 48bed5b8-b284-4b52-4143-153ae1c751a4
ms.date: 06/08/2017
---


# EncryptionProvider Members (Office)
Provides the methods for setting up permissions, applying the cryptography of the underlying encryption and decryption, and user authentication. 

Provides the methods for setting up permissions, applying the cryptography of the underlying encryption and decryption, and user authentication. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Authenticate](encryptionprovider-authenticate-method-office.md)|Used to determine whether the user has the proper permissions to open the encrypted document.|
|[CloneSession](encryptionprovider-clonesession-method-office.md)|Creates a second, working copy of the  **EncryptionProvider** object's encryption session for a file that is about to be saved.|
|[DecryptStream](encryptionprovider-decryptstream-method-office.md)|Decrypts and returns a stream of encrypted data for a document.|
|[EncryptStream](encryptionprovider-encryptstream-method-office.md)|Encrypts and returns a stream of data for a document.|
|[EndSession](encryptionprovider-endsession-method-office.md)|Ends the current encryption session.|
|[GetProviderDetail](encryptionprovider-getproviderdetail-method-office.md)|Displays information about the encryption of the current document. |
|[NewSession](encryptionprovider-newsession-method-office.md)|Used by the  **EncryptionProvider** object to create a new encryption session. This session is used by the provider to cache document-specific information about the encryption, users, and rights while the document is in memory.|
|[Save](encryptionprovider-save-method-office.md)|Saves an encrypted document.|
|[ShowSettings](encryptionprovider-showsettings-method-office.md)|Used to display a dialog of the encryption settings for the current document.|

