---
title: "Свойство BorderArtFormat.StretchPictures (издатель)"
keywords: vbapb10.chm7602181
f1_keywords: vbapb10.chm7602181
ms.prod: publisher
api_name: Publisher.BorderArtFormat.StretchPictures
ms.assetid: d3a9c867-111c-a4b1-0e56-6e5ed1e52c8c
ms.date: 06/08/2017
ms.openlocfilehash: 35d35626f4ccd271fd27cedff4ec057efd84069d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformatstretchpictures-property-publisher"></a>Свойство BorderArtFormat.StretchPictures (издатель)

 **Значение true,** чтобы увеличить картинка изображение, составляющих указанного Узорные в соответствии со фигуры, к которым применяется. Чтение и запись **типа Boolean**. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **StretchPictures**

 переменная _expression_A, представляет собой объект- **BorderArtFormat** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Возвращает «Отказано в разрешении», если не была применена Узорные на указанный объект.

Соответствует **не увеличить изображения** и **увеличить изображения в соответствии с** элементами управления в диалоговое окно **Узорные** .


## <a name="example"></a>Пример

Следующий пример проверяет наличие Узорные на каждой фигуры для каждой страницы активных документов. Если существует Узорные имеет значение, чтобы оно может быть растянуто.


```vb
Sub StretchBorderArt() 
 Dim anyPage As Page 
 Dim anyShape As Shape 
 
 For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .StretchPictures = True 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArtFormat](borderartformat-object-publisher.md)

