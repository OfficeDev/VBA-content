---
title: "Объект ColorsInUse (издатель)"
keywords: vbapb10.chm3014655
f1_keywords: vbapb10.chm3014655
ms.prod: publisher
api_name: Publisher.ColorsInUse
ms.assetid: ced0028a-8ab5-d9b1-b28c-24b794bdcbfe
ms.date: 06/08/2017
ms.openlocfilehash: 9d69a11b55a0d404f5a91e0d73d18ae3d35bdbf7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorsinuse-object-publisher"></a>Объект ColorsInUse (издатель)

Коллекция объектов **[ColorFormat](colorformat-object-publisher.md)** , представляющих цвета в указанной публикации.
 


## <a name="remarks"></a>Заметки

Коллекция **ColorsInUse** поддерживает все модели цвет публикации: RGB, обработки и плашечный цвет.
 

 
Для процесса цвет и публикации плашечные цвета основаны на красок. Для данного рукописного ввода публикации может содержать несколько цветов, которые являются разные оттенки и тени, рукописного ввода. Используйте коллекцию **[форм](plates-object-publisher.md)** для доступа к формы, которые представляют краски, определенной для публикации.
 

 

## <a name="example"></a>Пример

Используйте свойство **[ColorsInUse](http://msdn.microsoft.com/library/b018ffbc-b848-c0d0-19fa-df053e45260d%28Office.15%29.aspx)** объекта **[Document](document-object-publisher.md)** для возврата коллекции **ColorsInUse** .
 

 
В следующем примере перечисляются свойства каждого цвета в активной публикации, основанный на указанном рукописного ввода. В этом примере предполагается, что режим цвета публикации был определен в качестве плашечных или процесс и плашечных цветов.
 

 



```
Sub ListColorsBasedOnInk() 
Dim cfLoop As ColorFormat 
 
For Each cfLoop In ActiveDocument.ColorsInUse 
 
 With cfLoop 
 If .Ink = "2" Then 
 Debug.Print "BaseRGB: " &amp; .BaseRGB 
 Debug.Print "RGB: " &amp; .RGB 
 Debug.Print "TintShade: " &amp; .TintAndShade 
 Debug.Print "Type: " &amp; .Type 
 End If 
 End With 
 
Next cfLoop 
 
End Sub
```

Используйте **ColorsInUse** (индекс), где индекс — номер индекса цвета, чтобы получить объект **ColorFormat** . Следующий пример возвращает свойства для второй цвет в публикации.
 

 



```
Sub ColorProperties() 
 
 With ActiveDocument.ColorsInUse(2) 
 Debug.Print "Color RBG: " &amp; .RGB 
 Debug.Print "Ink RBG: " &amp; .BaseRGB 
 Debug.Print "Tint: " &amp; .TintAndShade 
 
 End With 
 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](colorsinuse-application-property-publisher.md)|
|[Count](colorsinuse-count-property-publisher.md)|
|[Элемент](colorsinuse-item-property-publisher.md)|
|[Родительский раздел](colorsinuse-parent-property-publisher.md)|

