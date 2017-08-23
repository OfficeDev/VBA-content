---
title: "Метод FillFormat.Patterned (издатель)"
keywords: vbapb10.chm2359314
f1_keywords: vbapb10.chm2359314
ms.prod: publisher
api_name: Publisher.FillFormat.Patterned
ms.assetid: 10e363b7-1160-55d3-5c97-733b7742b619
ms.date: 06/08/2017
ms.openlocfilehash: 7a9448bd8e7e5969cf8f34743c0e89dce4957865
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatpatterned-method-publisher"></a>Метод FillFormat.Patterned (издатель)

Устанавливает указанный заливки с шаблоном.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Узорные** ( **_Шаблон_**)

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Шаблон|Обязательное свойство.| **MsoPatternType**|Шаблон, используемый для указанного заливки.|

## <a name="remarks"></a>Заметки

Параметр шаблон может иметь одно из **MsoPatternType** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



| **msoPattern5Percent**|| **msoPattern10Percent**|| **msoPattern20Percent**|| **msoPattern25Percent**|| **msoPattern30Percent**|| **msoPattern40Percent**|| **msoPattern50Percent**|| **msoPattern60Percent**|| **msoPattern70Percent**|| **msoPattern75Percent**|| **msoPattern80Percent**|| **msoPattern90Percent**|| **msoPatternDarkDownwardDiagonal**|| **msoPatternDarkHorizontal**|| **msoPatternDarkUpwardDiagonal**|| **msoPatternDarkVertical**|| **msoPatternDashedDownwardDiagonal**|| **msoPatternDashedHorizontal**|| **msoPatternDashedUpwardDiagonal**|| **msoPatternDashedVertical**|| **msoPatternDiagonalBrick**|| **msoPatternDivot**|| **msoPatternDottedDiamond**|| **msoPatternDottedGrid**|| **msoPatternHorizontalBrick**|| **msoPatternLargeCheckerBoard**|| **msoPatternLargeConfetti**|| **msoPatternLargeGrid**|| **msoPatternLightDownwardDiagonal**|| **msoPatternLightHorizontal**|| **msoPatternLightUpwardDiagonal**|| **msoPatternLightVertical**|| **msoPatternNarrowHorizontal**|| **msoPatternNarrowVertical**|| **msoPatternOutlinedDiamond**|| **msoPatternPlaid**|| **msoPatternShingle**|| **msoPatternSmallCheckerBoard**|| **msoPatternSmallConfetti**|| **msoPatternSmallGrid**|| **msoPatternSolidDiamond**|| **msoPatternSphere**|| **msoPatternTrellis**|| **msoPatternWave**|| **msoPatternWeave**|| **msoPatternWideDownwardDiagonal**|| **msoPatternWideUpwardDiagonal**|| **msoPatternZigZag**| Используйте свойства [BackColor](fillformat-backcolor-property-publisher.md)и [ForeColor](fillformat-forecolor-property-publisher.md)Установка цвета, используемые в шаблоне.


## <a name="example"></a>Пример

В этом примере добавляется Овал с Узорная заливка active публикацию.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=60, Top:=60, Width:=80, Height:=40).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 0, 255) 
 .Patterned Pattern:=msoPatternDarkVertical 
End With 

```


