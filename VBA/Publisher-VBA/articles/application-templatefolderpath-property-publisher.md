---
title: "Свойство Application.TemplateFolderPath (издатель)"
keywords: vbapb10.chm131120
f1_keywords: vbapb10.chm131120
ms.prod: publisher
api_name: Publisher.Application.TemplateFolderPath
ms.assetid: e2256af9-9432-6205-864a-10bb7dec41c9
ms.date: 06/08/2017
ms.openlocfilehash: ed05f01a45ae7357876f96bafb0a8f4df48e7652
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationtemplatefolderpath-property-publisher"></a>Свойство Application.TemplateFolderPath (издатель)

Возвращает **строку** , представляющую расположение, где хранятся шаблоны Microsoft Publisher. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TemplateFolderPath**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере создается новая публикация и изменение главной страницы, содержащие номер страницы в звезда в левом верхнем углу страницы; Затем он сохраняет новые публикации местоположение папки шаблона, чтобы его можно использовать в качестве шаблона.


```vb
Sub CreateNewPubTemplate() 
 Dim AppPub As Application 
 Dim DocPub As Document 
 Dim strFolder As String 
 
 Set AppPub = New Publisher.Application 
 Set DocPub = AppPub.NewDocument 
 AppPub.ActiveWindow.Visible = True 
 strFolder = AppPub.TemplateFolderPath 
 
 With DocPub 
 With .MasterPages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=36, _ 
 Top:=36, Width:=50, Height:=50) 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 With .TextFrame.TextRange 
 .InsertPageNumber 
 .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
 With .Font 
 .Bold = msoTrue 
 .Color.RGB = RGB(Red:=255, Green:=255, Blue:=255) 
 .Size = 12 
 End With 
 End With 
 End With 
 .SaveAs FileName:=strFolder &; "\NewPubTemplt.pub" 
 End With 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

