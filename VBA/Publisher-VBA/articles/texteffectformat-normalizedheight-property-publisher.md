---
title: "Свойство TextEffectFormat.NormalizedHeight (издатель)"
keywords: vbapb10.chm3735814
f1_keywords: vbapb10.chm3735814
ms.prod: publisher
api_name: Publisher.TextEffectFormat.NormalizedHeight
ms.assetid: 2b62fe23-9204-7449-1d4e-73e73def5df0
ms.date: 06/08/2017
ms.openlocfilehash: 4273715ac9924a5d2b2e70ad84a766d0627429ee
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformatnormalizedheight-property-publisher"></a>Свойство TextEffectFormat.NormalizedHeight (издатель)

Указывает, являются ли все символы (прописные и строчные) в указанном WordArt по высоте. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **NormalizedHeight**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **NormalizedHeight** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Символы в указанном объекте WordArt не все же высоту.|
| **msoTrue**| Символы в указанном объекте WordArt — это все же высоту.|

## <a name="example"></a>Пример

В этом примере создается новая форма WordArt на первой странице active публикации и затем задает каждый символ в форму на одинаковую высоту.


```vb
Sub SetNormalHeight() 
 With ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect10, _ 
 text:="Test WordArt Shape", FontName:="Snap ITC", _ 
 FontSize:=30, FontBold:=msoFalse, FontItalic:=msoFalse, _ 
 Left:=36, Top:=36).TextEffect 
 .NormalizedHeight = msoTrue 
 End With 
End Sub
```


