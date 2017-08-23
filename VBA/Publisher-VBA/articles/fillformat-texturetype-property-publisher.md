---
title: "Свойство FillFormat.TextureType (издатель)"
keywords: vbapb10.chm2359568
f1_keywords: vbapb10.chm2359568
ms.prod: publisher
api_name: Publisher.FillFormat.TextureType
ms.assetid: 08f3b0a1-97a3-bdbf-25b4-93e05938d607
ms.date: 06/08/2017
ms.openlocfilehash: 4d9120f11f7e2f5231beaed521b25aa1b9a077e6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformattexturetype-property-publisher"></a>Свойство FillFormat.TextureType (издатель)

Возвращает константу **MsoTextureType** , указывающий тип текстуры для указанного заполнения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextureType**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTextureType


## <a name="remarks"></a>Заметки

Это свойство доступно только для чтения. Используйте метод [PresetTextured](fillformat-presettextured-method-publisher.md)или **[UserTextured](fillformat-usertextured-method-publisher.md)** для задания типа текстуры для заполнения.

Значение свойства может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoTexturePreset**| Заливки использует тип, предварительно текстуры.|
| **msoTextureTypeMixed**|Указывает оба типа текстуры для диапазона указанной фигуры.|
| **msoTextureUserDefined**|Заливки использует тип, определенный пользователем текстуры.|

## <a name="example"></a>Пример

В этом примере применяется полотно текстуры для заполнения для всех фигур на первой странице active публикации, которые в настоящий момент назначены заливки с пользовательской текстуры.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop.Fill 
 If .TextureType = msoTextureUserDefined Then 
 .PresetTextured _ 
 PresetTexture:=msoTextureCanvas 
 End If 
 End With 
Next shpLoop
```


