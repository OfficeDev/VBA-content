---
title: "Свойство Document.DefaultTabStop (издатель)"
keywords: vbapb10.chm196616
f1_keywords: vbapb10.chm196616
ms.prod: publisher
api_name: Publisher.Document.DefaultTabStop
ms.assetid: 245ff7a3-9828-5220-b692-2ce6effb9eb6
ms.date: 06/08/2017
ms.openlocfilehash: 6293670c1dd6cf4689c24d8d331b81974475a62a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentdefaulttabstop-property-publisher"></a>Свойство Document.DefaultTabStop (издатель)

Возвращает или задает **Variant** соответствующий позиции табуляции по умолчанию для всего текста в активной публикации. Допустимые значения — от 1 для 1584 точек (0.014" для 22"). После установки числовых значений считаются в пунктах. **Строковые** значения может находиться в любой единицы, поддерживаемый Microsoft Publisher. Точка значения всегда возвращаются. Если значения вне допустимого диапазона, возвращается ошибка. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DefaultTabStop**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Используйте метод **[InchesToPoints не была назначена](application-inchestopoints-method-publisher.md)** для преобразования дюймов в пунктах.


## <a name="example"></a>Пример

В этом примере задается свойство **DefaultTabStop** 72 точки для всего текста в активной публикации.


```vb
Sub SetTab() 
 Application.ActiveDocument.DefaultTabStop = 72 
End Sub 
```


