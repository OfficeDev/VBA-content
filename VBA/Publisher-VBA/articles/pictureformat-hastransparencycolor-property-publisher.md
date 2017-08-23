---
title: "Свойство PictureFormat.HasTransparencyColor (издатель)"
keywords: vbapb10.chm3604789
f1_keywords: vbapb10.chm3604789
ms.prod: publisher
api_name: Publisher.PictureFormat.HasTransparencyColor
ms.assetid: 2e6066e8-60b0-c33e-0bb0-1b6f83208fd0
ms.date: 06/08/2017
ms.openlocfilehash: 1455e1c07092c70251392019adcd146b8898e6b1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformathastransparencycolor-property-publisher"></a>Свойство PictureFormat.HasTransparencyColor (издатель)

Возвращает значение **типа Boolean** , которое указывает, применяется ли цвет прозрачность для указанного изображения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasTransparencyColor**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В следующем примере возвращается список изображений с прозрачность цвета в активной публикации.


```vb
Sub ListPicturesWithTransColors() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 If .HasTransparencyColor = True Then 
 Debug.Print .Filename 
 End If 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
 Next pgLoop 
 
End Sub
```


