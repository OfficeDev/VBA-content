---
title: "Свойство Font.Kerning (издатель)"
keywords: vbapb10.chm5373976
f1_keywords: vbapb10.chm5373976
ms.prod: publisher
api_name: Publisher.Font.Kerning
ms.assetid: 756fe3fa-9bf3-be16-2dd1-5b8fb0ec6496
ms.date: 06/08/2017
ms.openlocfilehash: 8d7a68a678424ce72736d793f2a28f7c570b2190
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontkerning-property-publisher"></a>Свойство Font.Kerning (издатель)

Возвращает или задает **Variant** , указывающее количество интервал по горизонтали, Microsoft Publisher применяется к символов в диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Кернинг**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Когда для этого свойства числовых значений считаются в пунктах и **строковые** значения может находиться в любой единицы поддерживаются в Publisher. Возвращаемые значения: **одного** типа и в пунктах. Отрицательные значения Объединить знаки ближе чем обычно и положительные значения распространение символов дальше друг от друга, чем обычно. Допустимые значения — от-600.0 к 600.0 точек.

Используйте метод **[InchesToPoints не была назначена](application-inchestopoints-method-publisher.md)** для преобразования дюймов в пунктах.


## <a name="example"></a>Пример

Этот пример устанавливает кернинг весь текст в первой статьи 6 момент.


```vb
Application.ActiveDocument.Stories(1).TextRange.Font.Kerning = 6 

```


