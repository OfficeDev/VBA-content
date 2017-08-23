---
title: "Метод FillFormat.PresetTextured (издатель)"
keywords: vbapb10.chm2359316
f1_keywords: vbapb10.chm2359316
ms.prod: publisher
api_name: Publisher.FillFormat.PresetTextured
ms.assetid: 971eac34-4e29-c898-93c8-9e71bd92238d
ms.date: 06/08/2017
ms.openlocfilehash: 1f837712b6d9cc31acbe7905c270c6bcdb6082a4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatpresettextured-method-publisher"></a>Метод FillFormat.PresetTextured (издатель)

Устанавливает указанный заливки предварительно текстуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetTextured** ( **_PresetTexture_**)

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PresetTexture|Обязательное свойство.| **MsoPresetTexture**|Предварительно текстуры.|

## <a name="remarks"></a>Заметки

Параметр PresetTexture может иметь одно из следующих **MsoPresetTexture** константы, описанные в библиотеке типов, Microsoft Office.



| **msoTextureBlueTissuePaper**|| **msoTextureBouquet**|| **msoTextureBrownMarble**|| **msoTextureCanvas**|| **msoTextureCork**|| **msoTextureDenim**|| **msoTextureFishFossil**|| **msoTextureGranite**|| **msoTextureGreenMarble**|| **msoTextureMediumWood**|| **msoTextureNewsprint**|| **msoTextureOak**|| **msoTexturePaperBag**|| **msoTexturePapyrus**|| **msoTextureParchment**|| **msoTexturePinkTissuePaper**|| **msoTexturePurpleMesh**|| **msoTextureRecycledPaper**|| **msoTextureSand**|| **msoTextureStationery**|| **msoTextureWalnut**|| **msoTextureWaterDroplets**|| **msoTextureWhiteMarble**|| **msoTextureWovenMat**|

## <a name="example"></a>Пример

В этом примере добавляется прямоугольник с заливкой текстуры зеленый мрамор active публикацию.


```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCan, _ 
 Left:=90, Top:=90, Width:=40, Height:=80) _ 
 .Fill.PresetTextured _ 
 PresetTexture:=msoTextureGreenMarble 

```


