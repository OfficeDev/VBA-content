---
title: OlExchangeConnectionMode Enumeration (Outlook)
keywords: vbaol11.chm3098
f1_keywords:
- vbaol11.chm3098
ms.prod: outlook
api_name:
- Outlook.OlExchangeConnectionMode
ms.assetid: ab43999d-f578-65ab-1f3d-455c66022901
ms.date: 06/08/2017
---


# OlExchangeConnectionMode Enumeration (Outlook)

Specifies whether the account is connected to an Exchange server and if so, the connection mode.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olCachedConnectedDrizzle**|600|The account is using cached Exchange mode such that headers are downloaded first, followed by the bodies and attachments of full items.|
| **olCachedConnectedFull**|700|The account is using cached Exchange mode on a Local Area Network or a fast connection with the Exchange server. The user can also select this state manually, disabling auto-detect logic and always downloading full items regardless of connection speed.|
| **olCachedConnectedHeaders**|500|The account is using cached Exchange mode on a dial-up or slow connection with the Exchange server, such that only headers are downloaded. Full item bodies and attachments remain on the server. The user can also select this state manually regardless of connection speed.|
| **olCachedDisconnected**|400|The account is using cached Exchange mode with a disconnected connection to the Exchange server.|
| **olCachedOffline**|200|The account is using cached Exchange mode and the user has selected  **Work Offline** from the **File** menu.|
| **olDisconnected**|300|The account has a disconnected connection to the Exchange server.|
| **olNoExchange**|0|The account does not use an Exchange server.|
| **olOffline**|100|The account is not connected to an Exchange server and is in the classic offline mode. This also occurs when the user selects  **Work Offline** from the **File** menu.|
| **olOnline**|800|The account is connected to an Exchange server and is in the classic online mode.|

