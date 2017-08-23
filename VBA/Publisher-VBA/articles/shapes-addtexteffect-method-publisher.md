---
title: "Метод Shapes.AddTextEffect (издатель)"
keywords: vbapb10.chm2162721
f1_keywords: vbapb10.chm2162721
ms.prod: publisher
api_name: Publisher.Shapes.AddTextEffect
ms.assetid: 21af82f1-d507-3c16-72df-bde1b5e00717
ms.date: 06/08/2017
ms.openlocfilehash: acc360ff465d9813eda8b8ddca5cab124467a254
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddtexteffect-method-publisher"></a>Метод Shapes.AddTextEffect (издатель)

Добавление нового объекта **Shape** , представляющий объект WordArt определенной коллекции **фигур** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddTextEffect** ( **_PresetTextEffect_**, **_текст_**, **_FontName_**, **_FontSize_**, **_FontBold_**, **_FontItalic_**, **_слева_**, **_в начало_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PresetTextEffect|Обязательное свойство.| **MsoPresetTextEffect**|Влияние предварительно текст для использования. Значения константы **MsoPresetTextEffect** соответствуют форматы, перечисленные в диалоговом окне **Коллекция WordArt** (нумерованные слева направо и сверху вниз).|
|Text|Обязательное свойство.| **String**|Текст, который используется для объекта WordArt.|
|FontName|Обязательное свойство.| **String**|Имя шрифта, используемого для объекта WordArt.|
|FontSize|Обязательное свойство.| **Variant**|Размер шрифта для объекта WordArt. Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).|
|FontBold|Обязательное свойство.| **MsoTriState**|Определяет, следует ли для форматирования текста WordArt как полужирным шрифтом.|
|FontItalic|Обязательное свойство.| **MsoTriState**|Определяет, следует ли для форматирования текста WordArt как курсив.|
|Слева|Обязательное свойство.| **Variant**|Положение левого края фигуры, представляющий объект WordArt.|
|Вверх|Обязательное свойство.| **Variant**|Положение верхнего края фигуры, представляющий объект WordArt.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Для параметров Left и Top числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Publisher (например, «2,5 дюйма»).

Высота и ширина объекта WordArt определяется его текста и форматирования.

Свойство **[TextEffect](shape-texteffect-property-publisher.md)** используется для возврата объекта **[TextEffectFormat](texteffectformat-object-publisher.md)** , свойства которого может использоваться для изменения существующего объекта WordArt.

Параметр PresetTextEffect может иметь одно из ** [MsoPresetTextEffect](http://msdn.microsoft.com/library/56a7008d-ce2c-f127-56de-851cb8fef44f%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office. Константа **msoTextEffectMixed** не поддерживается.

Параметр FontBold может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Не форматировать текст WordArt в качестве полужирным шрифтом.|
| **msoTrue**|Отформатируйте текст WordArt в качестве полужирным шрифтом.|
Параметр FontItalic может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Не формат текста WordArt курсив.|
| **msoTrue**|Формат текста WordArt курсив.|

## <a name="example"></a>Пример

Следующий пример добавляет объект WordArt первой страницы active публикации.


```vb
Dim shpWordArt As Shape 
 
Set shpWordArt = ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect7, Text:="Annual Report", _ 
 FontName:="Arial Black", FontSize:=24, _ 
 FontBold:=msoFalse, FontItalic:=msoFalse, _ 
 Left:=144, Top:=72) 

```


