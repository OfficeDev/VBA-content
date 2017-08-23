---
title: "Метод ThreeDFormat.SetThreeDFormat (издатель)"
keywords: vbapb10.chm3801107
f1_keywords: vbapb10.chm3801107
ms.prod: publisher
api_name: Publisher.ThreeDFormat.SetThreeDFormat
ms.assetid: d73dbada-1a33-4b5b-9733-e228a0cc5f8c
ms.date: 06/08/2017
ms.openlocfilehash: 47bc61763b1320332a4703e639ab1259948d8df3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatsetthreedformat-method-publisher"></a>Метод ThreeDFormat.SetThreeDFormat (издатель)

Задает формат предварительно придания объема. Каждый формат предварительно придания объема содержит набор предварительно значений для объемных свойства изменяется.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetThreeDFormat** ( **_PresetThreeDFormat_**)

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PresetThreeDFormat|Обязательное свойство.| **MsoPresetThreeDFormat**|Задает формат предварительно придания объема, соответствующее одному из параметров (нумерованные слева направо, сверху вниз) отображаются при нажатии кнопки **объемных** на панели инструментов **Рисование** .|

## <a name="remarks"></a>Заметки

Этот метод задает свойство **[PresetThreeDFormat](threedformat-presetthreedformat-property-publisher.md)** в формат, указанный в аргументе PresetThreeDFormat.

Параметр PresetThreeDFormat может быть одной из констант **MsoPresetThreeDFormat** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



| **msoThreeD1**|| **msoThreeD2**|| **msoThreeD3**|| **msoThreeD4**|| **msoThreeD5**|| **msoThreeD6**|| **msoThreeD7**|| **msoThreeD8**|| **msoThreeD9**|| **msoThreeD10**|| **msoThreeD11**|| **msoThreeD12**|| **msoThreeD13**|| **msoThreeD14**|| **msoThreeD15**|| **msoThreeD16**|| **msoThreeD17**|| **msoThreeD18**|| **msoThreeD19**|| **msoThreeD20**|

## <a name="example"></a>Пример

В этом примере добавляется овала active публикацию и задает его формате придания объема к одной из предварительно трехмерной форматы.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=30, Top:=30, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .SetThreeDFormat PresetThreeDFormat:=msoThreeD12 
End With 

```


