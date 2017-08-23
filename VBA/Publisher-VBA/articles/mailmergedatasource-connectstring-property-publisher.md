---
title: "Свойство MailMergeDataSource.ConnectString (издатель)"
keywords: vbapb10.chm6291460
f1_keywords: vbapb10.chm6291460
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.ConnectString
ms.assetid: d7719567-f946-6b76-3ff2-d372dcc76a17
ms.date: 06/08/2017
ms.openlocfilehash: 6742c8baf290a2bae3e9e0dd41540f36cc5e5de6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceconnectstring-property-publisher"></a>Свойство MailMergeDataSource.ConnectString (издатель)

Возвращает **строку** , представляющую подключения для указанного источника данных. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Стрсоедин**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере проверяется, если строка подключения содержит указанные символы OLEDB и отображает сообщение соответствующим образом.


```vb
Sub VerifyCorrectDataSource() 
 
 With ActiveDocument.MailMerge.DataSource 
 If InStr(.ConnectString, "OLEDB") > 0 Then 
 MsgBox "OLE DB is used to connect to the data source." 
 Else 
 MsgBox "OLE DB is not used to connect to the data source." 
 End If 
 End With 
 
End Sub
```


