---
title: "Свойство Application.Version (издатель)"
keywords: vbapb10.chm131121
f1_keywords: vbapb10.chm131121
ms.prod: publisher
api_name: Publisher.Application.Version
ms.assetid: ffec5bca-cd81-77c6-d80b-e629abfa6dec
ms.date: 06/08/2017
ms.openlocfilehash: 680f5603ea66f5b8d3cabfb10d643f5b8cf29683
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationversion-property-publisher"></a>Свойство Application.Version (издатель)

Возвращает **строку** , указывающую номер версии установленного копию Microsoft Publisher. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Версия**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

Следующий пример номер версии и построения-установленные копии Publisher.


```vb
MsgBox "You are currently running Microsoft Publisher, " _ 
 &; " version " &; Application.Version &; ", build " _ 
 &; Application.Build &; "." 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

