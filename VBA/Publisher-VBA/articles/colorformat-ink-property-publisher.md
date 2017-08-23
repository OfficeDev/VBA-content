---
title: "Свойство ColorFormat.Ink (издатель)"
keywords: vbapb10.chm2555911
f1_keywords: vbapb10.chm2555911
ms.prod: publisher
api_name: Publisher.ColorFormat.Ink
ms.assetid: 53851337-fdce-7b72-5626-50bce370457b
ms.date: 06/08/2017
ms.openlocfilehash: 3c55b43d2ef935cc39cb2575026390aa64a57de9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorformatink-property-publisher"></a>Свойство ColorFormat.Ink (издатель)

Возвращает или задает **Long** , указывающее, является ли указанный цвет является плашечным, и если да, то место печатных которой он принадлежит. Допустимые значения: **pbInkNone** (значение по умолчанию; значение, то, что цвет не плашечный цвет) или число в интервале между 1 и _n_ где _n_ — это число формы смесевых цветов. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Рукописного ввода**

 переменная _expression_A, представляющий объект **ColorFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

Следующий пример указывает, что цвет первый диапазон текста на странице один активный публикации должна быть назначена плашечных форме двух.


```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Font.Color.Ink = 2
```


