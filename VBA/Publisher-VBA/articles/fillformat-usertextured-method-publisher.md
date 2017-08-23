---
title: "Метод FillFormat.UserTextured (издатель)"
keywords: vbapb10.chm2359320
f1_keywords: vbapb10.chm2359320
ms.prod: publisher
api_name: Publisher.FillFormat.UserTextured
ms.assetid: fe1a1e06-8bdc-8022-6d4b-6f320f587baf
ms.date: 06/08/2017
ms.openlocfilehash: 2df074109249847ff9add2b0bdfae0882c0805a8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatusertextured-method-publisher"></a>Метод FillFormat.UserTextured (издатель)

Заполняет указанный фигуры малых заголовков изображения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UserTextured** ( **_TextureFile_**)

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|TextureFile|Обязательное свойство.| **String**|Имя файла текстуры.|

## <a name="remarks"></a>Заметки

Чтобы заполнить форму одно большое изображение, используйте метод **[UserPicture](fillformat-userpicture-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере добавляется два прямоугольника active публикацию. Область в левой части заполняется одно большое изображение с; прямоугольник справа заполняется множество небольших заголовков и то же изображение. (Обратите внимание на то, что действительный путь к файлу для работы этого примера необходимо заменить PathToFile.)


```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserPicture PictureFile:="PathToFile" 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserTextured TextureFile:="PathToFile" 
End With 

```


