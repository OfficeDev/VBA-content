---
title: "Свойство ColorSchemes.Item (издатель)"
keywords: vbapb10.chm2752512
f1_keywords: vbapb10.chm2752512
ms.prod: publisher
api_name: Publisher.ColorSchemes.Item
ms.assetid: 5a66a0ae-b552-0979-d3ac-7b1d7bec96f7
ms.date: 06/08/2017
ms.openlocfilehash: bdd27da5580d14873f870404d6986650fedf4d25
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorschemesitem-property-publisher"></a>Свойство ColorSchemes.Item (издатель)

Возвращает указанный объект **[ColorScheme](colorscheme-object-publisher.md)** **ColorSchemes** коллекции. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **ColorSchemes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Variant**|Цветовая схема для возврата. Может быть либо именем цветовая схема как строку или соответствующей константы **PbColorScheme** .|

## <a name="remarks"></a>Заметки

Значение свойства **элемента** может иметь одно из **[PbColorScheme](pbcolorscheme-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере задается цветовая схема active публикации цветовая схема голубой.


```vb
ActiveDocument.ColorScheme = ColorSchemes.Item(Index:=pbColorSchemeAqua)
```


