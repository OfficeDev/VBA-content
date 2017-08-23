---
title: "Свойство Application.Options (издатель)"
keywords: vbapb10.chm131095
f1_keywords: vbapb10.chm131095
ms.prod: publisher
api_name: Publisher.Application.Options
ms.assetid: 999f208a-02e6-49fb-c9a0-42aa97c5e37e
ms.date: 06/08/2017
ms.openlocfilehash: 56024bfb9f6aca77804694e37b78a107ab5992bd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationoptions-property-publisher"></a>Свойство Application.Options (издатель)

Возвращает объект, представляющий параметры приложения, которые можно задать в Microsoft Publisher на **[Параметры](options-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Параметры**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Параметры


## <a name="example"></a>Пример

В этом примере показано отключение фоновое сохранение, а затем сохраняет active публикации.


```vb
Sub SetGlobalSaveOptions() 
 
 With Options 
 .AllowBackgroundSave = False 
 End With 
 
 ActiveDocument.Save 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

