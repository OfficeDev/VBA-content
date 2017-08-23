---
title: "Свойство Application.WebOptions (издатель)"
keywords: vbapb10.chm131176
f1_keywords: vbapb10.chm131176
ms.prod: publisher
api_name: Publisher.Application.WebOptions
ms.assetid: 2e0c3435-a55a-4903-a0f8-9c347dec03b5
ms.date: 06/08/2017
ms.openlocfilehash: 928940906f58574f28acc0e66bb12d59f8290350
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationweboptions-property-publisher"></a>Свойство Application.WebOptions (издатель)

Возвращает объект **[WebOptions](weboptions-object-publisher.md)** , который представляет свойства веб-публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebOptions**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

WebOptions


## <a name="example"></a>Пример

Следующий пример указывает, что веб-публикации не всегда следует хранить в используемых по умолчанию кодировки, и что кодировки должен быть Юникод (UTF-8).


```vb
With Application.WebOptions 
 .AlwaysSaveInDefaultEncoding = False 
 .Encoding = msoEncodingUTF8 
End With
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

