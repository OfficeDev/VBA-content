---
title: "Свойство ColorScheme.Name (издатель)"
keywords: vbapb10.chm2686979
f1_keywords: vbapb10.chm2686979
ms.prod: publisher
api_name: Publisher.ColorScheme.Name
ms.assetid: 8816c7d5-6dac-f1ad-f7f7-590406be5bef
ms.date: 06/08/2017
ms.openlocfilehash: 7253783b552c3e223f653e79b5a22f6c607e4fe1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorschemename-property-publisher"></a>Свойство ColorScheme.Name (издатель)

Возвращает **строковое** значение, указывающее имя указанного объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляет собой объект- **ColorScheme** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


## <a name="example"></a>Пример

В этом примере приводится имя цветовая схема active публикации.


```vb
MsgBox "The current color scheme is " _ 
 &; ActiveDocument.ColorScheme.Name &; "."
```


