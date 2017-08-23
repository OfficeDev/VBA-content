---
title: "Объект BorderArtFormat (издатель)"
keywords: vbapb10.chm7667711
f1_keywords: vbapb10.chm7667711
ms.prod: publisher
api_name: Publisher.BorderArtFormat
ms.assetid: ba066b2e-fe40-aeef-9275-2cc2810f63ca
ms.date: 06/08/2017
ms.openlocfilehash: 01793f52b174cd8dc494ddc3d858a8acbe266a3c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformat-object-publisher"></a>Объект BorderArtFormat (издатель)

Представляет форматирования Узорные, применяемые к указанной фигуры.
 


## <a name="remarks"></a>Заметки

Узорные, границы изображения, которые можно применять для текстовых полей, рамки рисунков или прямоугольники.
 

 

## <a name="example"></a>Пример

Свойство **[Узорные](shape-borderart-property-publisher.md)** фигуры возвращает объект **BorderArtFormat** .
 

 
Следующий пример возвращает Узорные первую фигуру на первой странице active публикации и отображает имя Узорные в окне сообщения.
 

 



```
Dim bdaTemp As BorderArtFormat 
 
Set bdaTemp = ActiveDocument.Pages(1).Shapes(1).BorderArt 
MsgBox "BorderArt name is: " &amp;bdaTemp.Name
```

Позволяет указать тип, который требуется Узорные применяется к изображению метод **[Set](borderartformat-set-method-publisher.md)** . Следующий пример проверяет наличие Узорные на каждой фигуры для каждой страницы активных документов. Все найденные Узорные присвоено значение того же типа.
 

 



```
Sub SetBorderArt() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Set(strBorderArtName) 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```

Свойство **[Name](borderartformat-name-property-publisher.md)** можно также использовать для указания того, какой тип Узорные требуется применяется к изображению. В следующем примере задается Узорные в документе для того же типа с помощью свойства **Name** .
 

 



```
Sub SetBorderArtByName() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Name = strBorderArtName 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```


 **Примечание**  Так как **имя** является свойством по умолчанию как **[Узорные](borderart-object-publisher.md)** , так и **BorderArtFormat** объектов, его состояния явным образом, при задании типа Узорные не требуется. Оператор `Shape.BorderArtFormat = Document.BorderArts(1)`соответствует`Shape.BorderArtFormat.Name = Document.BorderArts(1).Name`
 

Используйте метод **[Delete](borderartformat-delete-method-publisher.md)** для удаления Узорные из изображения. Следующий пример проверяет существование картинка границы на каждой фигуры для каждой страницы активных документов. Если картинка границы существует, она удаляется.
 

 



```
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


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](borderartformat-delete-method-publisher.md)|
|[RevertToDefaultWeight](borderartformat-reverttodefaultweight-method-publisher.md)|
|[RevertToOriginalColor](borderartformat-reverttooriginalcolor-method-publisher.md)|
|[SET](borderartformat-set-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](borderartformat-application-property-publisher.md)|
|[Цвет](borderartformat-color-property-publisher.md)|
|[Существует](borderartformat-exists-property-publisher.md)|
|[Name](borderartformat-name-property-publisher.md)|
|[Родительский раздел](borderartformat-parent-property-publisher.md)|
|[StretchPictures](borderartformat-stretchpictures-property-publisher.md)|
|[Вес](borderartformat-weight-property-publisher.md)|

