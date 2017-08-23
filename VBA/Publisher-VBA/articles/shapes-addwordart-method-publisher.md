---
title: "Метод Shapes.AddWordArt (издатель)"
keywords: vbapb10.chm2162761
f1_keywords: vbapb10.chm2162761
ms.prod: publisher
api_name: Publisher.Shapes.AddWordArt
ms.assetid: 8ff83baa-5d88-5f80-3a69-5f712ba5e583
ms.date: 06/08/2017
ms.openlocfilehash: 1b6d2aaa8051cdb1881086e626cbc4281e3fad94
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddwordart-method-publisher"></a>Метод Shapes.AddWordArt (издатель)

Возвращает объект **фигуры** , представляющий WordArt для добавления к публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddWordArt** ( **_PresetWordArt_**, **_текст_**, **_FontName_**, **_FontSize_**, **_FontBold_**, **_FontItalic_**, **_слева_**, **_в начало_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PresetWordArt|Обязательное свойство.| **pbPresetWordArt**|Тип предварительно WordArt для добавления.|
|Text|Обязательное свойство.| **String**|Текст WordArt.|
|FontName|Обязательное свойство.| **String**|Имя шрифта для использования в объект WordArt.|
|FontSize|Обязательное свойство.| **Variant**|Размер шрифта для использования в объект WordArt.|
|FontBold|Обязательное свойство.| **[MSOTRISTATE]**|Текст WordArt должны ли полужирным шрифтом. Возможные значения см.|
|FontItalic|Обязательное свойство.| **[MSOTRISTATE]**|Текст WordArt должны ли курсив. Возможные значения см.|
|Слева|Обязательное свойство.| **Variant**|Горизонтальную позицию WordArt.|
|Вверх|Обязательное свойство.| **Variant**|Вертикальное положение WordArt.|

### <a name="return-value"></a>Возвращаемое значение

 **Фигура**


### <a name="remarks"></a>Заметки

Значение параметра **FontBold** может иметь одно из следующих **MsoTriState** константы, описанные в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в объект WordArt форматируются полужирным шрифтом.|
| **msoTriStateMixed**|Возвращает значение, указывающее, что объект WordArt содержит текст полужирным и не форматированный текст полужирным шрифтом.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в объект WordArt форматируются полужирным шрифтом.|
Значение параметра **FontItalic** может иметь одно из следующих **MsoTriState** константы, описанные в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в объект WordArt форматируются как курсив.|
| **msoTriStateMixed**|Возвращает значение, указывающее, что объект WordArt содержит текст в формате курсив и текст не в формате курсив.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в объект WordArt форматируются как курсив.|

