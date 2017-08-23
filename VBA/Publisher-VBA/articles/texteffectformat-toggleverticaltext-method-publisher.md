---
title: "Метод TextEffectFormat.ToggleVerticalText (издатель)"
keywords: vbapb10.chm3735568
f1_keywords: vbapb10.chm3735568
ms.prod: publisher
api_name: Publisher.TextEffectFormat.ToggleVerticalText
ms.assetid: 627ddbcc-5951-70c6-4e54-de0e9a4bebec
ms.date: 06/08/2017
ms.openlocfilehash: a5a3287da83fb2d9e336622ba910e1eda118caff
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformattoggleverticaltext-method-publisher"></a>Метод TextEffectFormat.ToggleVerticalText (издатель)

Переключение потока текст в указанном WordArt горизонтальную по вертикали или наоборот.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ToggleVerticalText**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


## <a name="remarks"></a>Заметки

С помощью метода **ToggleVerticalText** меняет местами значения свойств **[Left](shape-left-property-publisher.md)** и **[Top](shape-top-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** , который представляет объект WordArt и оставляет свойства **[ширины](shape-width-property-publisher.md)** и **[высоты](shape-height-property-publisher.md)** без изменений.

Метод **[Отразить](shape-flip-method-publisher.md)** и **[Вращение](shape-rotation-property-publisher.md)** свойство объекта **[Shape](shape-object-publisher.md)** **[RotatedChars](texteffectformat-rotatedchars-property-publisher.md)** свойство и метод **ToggleVerticalText** объекта **[TextEffectFormat](texteffectformat-object-publisher.md)** все влияет на ориентация символов и направление потока текста в объект **фигуры** , представляющий WordArt. Может потребоваться проверить, узнайте, как объединить эффекты этих свойств и методов для получения результатов, который будет.


## <a name="example"></a>Пример

В этом примере добавляется объект WordArt, который содержит текст «Test» для активной публикации и коммутаторы из потока горизонтальный текст (по умолчанию для указанного стилей WordArt, **msoTextEffect1**) на вертикальное.


```vb
Dim shpTextEffect As Shape 
 
Set shpTextEffect = ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=100, Top:=100) 
 
shpTextEffect.TextEffect.ToggleVerticalText
```


