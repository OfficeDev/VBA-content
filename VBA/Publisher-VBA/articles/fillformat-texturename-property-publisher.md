---
title: "Свойство FillFormat.TextureName (издатель)"
keywords: vbapb10.chm2359561
f1_keywords: vbapb10.chm2359561
ms.prod: publisher
api_name: Publisher.FillFormat.TextureName
ms.assetid: 237a85ff-018d-f6b7-e94b-32e85fce65ab
ms.date: 06/08/2017
ms.openlocfilehash: 7ba95b1675a501b21dbea0e699906475baa58e86
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformattexturename-property-publisher"></a>Свойство FillFormat.TextureName (издатель)

Возвращает **строку** , указывающую имя файла текстуры для указанного заполнения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextureName**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Используйте метод **[UserTextured](fillformat-usertextured-method-publisher.md)** для установки файлов текстуры для заполнения.


## <a name="example"></a>Пример

В этом примере добавляется овала active публикацию. Если фигуры одно active публикацией заливки с помощью пользовательских текстуры, новые Овал будут иметь же заливки как фигуры один. Если фигуры один любой другой тип заливки, новые Овал будут иметь зеленый мрамора заливки.


```vb
Dim ffNew As FillFormat 
 
With ActiveDocument.Pages(1).Shapes 
 Set ffNew = .AddShape(Type:=msoShapeOval, _ 
 Left:=0, Top:=0, Width:=200, Height:=90).Fill 
 
 With .Item(1).Fill 
 If .Type = msoFillTextured And _ 
 .TextureType = msoTextureUserDefined Then 
 ffNew.UserTextured _ 
 TextureFile:=.TextureName 
 Else 
 ffNew.PresetTextured _ 
 PresetTexture:=msoTextureGreenMarble 
 End If 
 End With 
End With 

```


