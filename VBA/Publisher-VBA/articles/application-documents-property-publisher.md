---
title: "Свойство Application.Documents (издатель)"
keywords: vbapb10.chm131174
f1_keywords: vbapb10.chm131174
ms.prod: publisher
api_name: Publisher.Application.Documents
ms.assetid: dd48d68f-a6ae-b5c0-2a85-90abff1e6c5a
ms.date: 06/08/2017
ms.openlocfilehash: e1f4015814476ab49d755bd8519075a28dc27596
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationdocuments-property-publisher"></a>Свойство Application.Documents (издатель)

Возвращает коллекцию **[документов](documents-object-publisher.md)** , представляющий все открытые публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Документы**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Документы


## <a name="example"></a>Пример

В следующем примере перечисляются все открытые публикации.


```vb
Dim objDocument As Document 
Dim strMsg As String 
For Each objDocument In Documents 
 strMsg = strMsg &; objDocument.Name &; vbCrLf 
Next objDocument 
MsgBox Prompt:=strMsg, Title:="Current Documents Open", Buttons:=vbOKOnly
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

