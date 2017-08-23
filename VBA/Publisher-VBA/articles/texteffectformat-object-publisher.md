---
title: "Объект TextEffectFormat (издатель)"
keywords: vbapb10.chm3801087
f1_keywords: vbapb10.chm3801087
ms.prod: publisher
api_name: Publisher.TextEffectFormat
ms.assetid: 672d0ef0-cbcd-05ef-9aa5-b986c7b045ac
ms.date: 06/08/2017
ms.openlocfilehash: b75c4445364d36d70cc59f1d02e93010f1fb907b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformat-object-publisher"></a>Объект TextEffectFormat (издатель)

Содержит свойства и методы, которые применяются к объекты WordArt.
 


## <a name="example"></a>Пример

Свойство **TextEffect** используется для возврата объекта **TextEffectFormat** . В следующем примере задается имя шрифта и форматирования для фигуры одно на первой странице active публикации. В данном примере для работы фигуры один должен быть объектом WordArt.
 

 

```
Sub FormatWordArt() 
 With ActiveDocument.Pages(1).Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = MsoTrue 
 .FontItalic = MsoTrue 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[ToggleVerticalText](texteffectformat-toggleverticaltext-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Выравнивание](texteffectformat-alignment-property-publisher.md)|
|[Приложения](texteffectformat-application-property-publisher.md)|
|[FontBold](texteffectformat-fontbold-property-publisher.md)|
|[FontItalic](texteffectformat-fontitalic-property-publisher.md)|
|[FontName](texteffectformat-fontname-property-publisher.md)|
|[FontSize](texteffectformat-fontsize-property-publisher.md)|
|[KernedPairs](texteffectformat-kernedpairs-property-publisher.md)|
|[NormalizedHeight](texteffectformat-normalizedheight-property-publisher.md)|
|[Родительский раздел](texteffectformat-parent-property-publisher.md)|
|[PresetShape](texteffectformat-presetshape-property-publisher.md)|
|[PresetTextEffect](texteffectformat-presettexteffect-property-publisher.md)|
|[PresetWordArt](texteffectformat-presetwordart-property-publisher.md)|
|[RotatedChars](texteffectformat-rotatedchars-property-publisher.md)|
|Да|
|[Отслеживание](texteffectformat-tracking-property-publisher.md)|

