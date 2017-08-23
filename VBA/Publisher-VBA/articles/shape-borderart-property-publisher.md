---
title: "Свойство Shape.BorderArt (издатель)"
keywords: vbapb10.chm5308675
f1_keywords: vbapb10.chm5308675
ms.prod: publisher
api_name: Publisher.Shape.BorderArt
ms.assetid: dcc0ceb4-ef69-ffd3-e510-13dcb8d06832
ms.date: 06/08/2017
ms.openlocfilehash: cb1d3fee5dd07ce44a22d15d4e671d3363f334f5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeborderart-property-publisher"></a>Свойство Shape.BorderArt (издатель)

Возвращает объект **[BorderArtFormat](borderartformat-object-publisher.md)** , представляющий тип Узорные, применяемые к указанной фигуры. Возвращает «Отказано в разрешении», если Узорные не была применена к фигуре. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Узорные**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

BorderArtFormat


## <a name="remarks"></a>Заметки

Узорные, границы изображения, которые можно применять для текстовых полей, рамки рисунков или прямоугольники. 

Используйте свойство **Узорные** для применения, изменение и удаление Узорные из фигур в публикации.


## <a name="example"></a>Пример

Следующий пример проверяет наличие Узорные на каждой фигуры для каждой страницы публикации active. Если Узорные существует, она удаляется.


```vb
Sub DeleteBorderArt() 
 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .Delete 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```


