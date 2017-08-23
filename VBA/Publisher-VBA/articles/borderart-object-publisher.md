---
title: "Объект Узорные (издатель)"
keywords: vbapb10.chm7733247
f1_keywords: vbapb10.chm7733247
ms.prod: publisher
api_name: Publisher.BorderArt
ms.assetid: 464bec0f-7912-ab27-9593-7f1cb53da342
ms.date: 06/08/2017
ms.openlocfilehash: 47f8a6138300057ed6f480f5388c1dba33022b49
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderart-object-publisher"></a>Объект Узорные (издатель)

Представляет тип доступные Узорные. Узорные — границы изображения, которые можно применять для текстовых полей, рамки рисунков или прямоугольники. **Узорные** объект является элементом коллекции **[BorderArts](borderarts-object-publisher.md)** . Коллекция **BorderArts** содержит все Узорные, доступных для использования в указанной публикации.
 


## <a name="remarks"></a>Заметки

Коллекция **BorderArts** включает все пользовательские типы Узорные, создаваемые пользователем для указанной публикации.
 

 

## <a name="example"></a>Пример

Используйте свойство **[Item](borderarts-item-method-publisher.md)** коллекции **BorderArts** для получения определенного объекта Узорные. Аргумент Index свойство **Item** может быть номер или имя объекта Узорные.
 

 
В этом примере возвращается Узорные «Apples» из активной публикации. 
 

 



```
Dim bdaTemp As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts.Item (Index:="Apples") 
```

Позволяет указать тип, который требуется Узорные применяется к изображению свойства **[Name](borderart-name-property-publisher.md)** . В следующем примере задается Узорные в документе для того же типа с помощью свойства **Name** .
 

 



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


 

 

 **Примечание**  Так как **имя** является свойством по умолчанию объекта **Узорные** и объект **BorderArtFormat** , его состояния явным образом, при задании типа Узорные не требуется. Оператор `Shape.BorderArtFormat = Document.BorderArts(1)`соответствует`Shape.BorderArtFormat.Name = Document.BorderArts(1).Name`
 


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](borderart-application-property-publisher.md)|
|[Name](borderart-name-property-publisher.md)|
|[Родительский раздел](borderart-parent-property-publisher.md)|

