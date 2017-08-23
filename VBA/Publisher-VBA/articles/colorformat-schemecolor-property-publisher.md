---
title: "Свойство ColorFormat.SchemeColor (издатель)"
keywords: vbapb10.chm2555910
f1_keywords: vbapb10.chm2555910
ms.prod: publisher
api_name: Publisher.ColorFormat.SchemeColor
ms.assetid: 8b02c85c-a976-7b10-c4ea-6f881d702b55
ms.date: 06/08/2017
ms.openlocfilehash: 29b0fef46ea0e4ff7b9f11a8d595e3d507263154
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorformatschemecolor-property-publisher"></a>Свойство ColorFormat.SchemeColor (издатель)

Задает цвет текущего цветовая схема. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SchemeColor**

 переменная _expression_A, представляет собой объект- **ColorFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbSchemeColorIndex


## <a name="remarks"></a>Заметки

Значение свойства **SchemeColor** может иметь одно из **[PbSchemeColorIndex](pbschemecolorindex-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В следующем примере устанавливается цвет текста в форму одно на странице один активный публикации диакритических знаков цвета пять в текущей цветовой схеме.


```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Font.Color.SchemeColor =
```


```
pbSchemeColorAccent5
```


