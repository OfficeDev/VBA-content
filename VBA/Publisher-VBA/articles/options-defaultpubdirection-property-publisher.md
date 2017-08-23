---
title: "Свойство Options.DefaultPubDirection (издатель)"
keywords: vbapb10.chm1048624
f1_keywords: vbapb10.chm1048624
ms.prod: publisher
api_name: Publisher.Options.DefaultPubDirection
ms.assetid: 628352c1-040f-9ab1-d0f1-308b2c26679c
ms.date: 06/08/2017
ms.openlocfilehash: 0da5baa3b883318851ae4a88407b5b82af16920d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsdefaultpubdirection-property-publisher"></a>Свойство Options.DefaultPubDirection (издатель)

Возвращает или задает значение константы **PbDirectionType** , представляющий направление по умолчанию, в какой текст создается денежных средств при новой публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DefaultPubDirection**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

PbDirectionType


## <a name="remarks"></a>Заметки

Значение свойства **DefaultPubDirection** может иметь одно из **[PbDirectionType](pbdirectiontype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Это свойство приводит к ошибке, если не выполняется в бизнес-аналитики обмен сообщениями версии Microsoft Publisher (например, арабский).


## <a name="example"></a>Пример

В этом примере направление текста по умолчанию для новых публикаций и текста в бизнес-аналитики обмен сообщениями версии Publisher.


```vb
Sub SetDefaultDirection() 
 With Options 
 .DefaultPubDirection = pbDirectionRightToLeft 
 .DefaultTextFlowDirection = pbDirectionRightToLeft 
 End With 
End Sub
```


