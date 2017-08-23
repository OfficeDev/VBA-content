---
title: "Метод Application.NewDocument (издатель)"
keywords: vbapb10.chm131127
f1_keywords: vbapb10.chm131127
ms.prod: publisher
api_name: Publisher.Application.NewDocument
ms.assetid: 9beb6176-0c46-0ba0-8d41-a9021c624223
ms.date: 06/08/2017
ms.openlocfilehash: 0ef817876651d6211b5a75bf834356ddab706748
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationnewdocument-method-publisher"></a>Метод Application.NewDocument (издатель)

Возвращает объект **Document** , представляющий новую публикацию.


## <a name="syntax"></a>Синтаксис

 _выражение_. **NewDocument** ( **_Мастер_** **_разработки_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Мастер|Необязательный| **PbWizard**|Мастер для создания новой публикации.|
|Разработка|Необязательный| **Длинный**|Разработка, чтобы применить новые публикации.|

### <a name="return-value"></a>Возвращаемое значение

Документ


## <a name="remarks"></a>Заметки

Параметр мастера может быть одной из констант **PbWizard** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице. Значение по умолчанию — **pbWizardNone**.





| **pbWizardAdvertisements**|| **pbWizardAirplanes**|| **pbWizardBanners**|| **pbWizardBrochures**|| **pbWizardBusinessCards**|| **pbWizardBusinessForms**|| **pbWizardCalendars**|| **pbWizardCatalogs**|| **pbWizardCertificates**|| **pbWizardEnvelopes**|| **pbWizardFlyers**|| **pbWizardGiftCertificates**|| **pbWizardGreetingCards**|| **pbWizardInvitations**|| **pbWizardJapaneseAdvertisements**|| **pbWizardJapaneseAirplanes**|| **pbWizardJapaneseBanners**|| **pbWizardJapaneseBrochures**|| **pbWizardJapaneseBusinessCards**|| **pbWizardJapaneseBusinessForms**|| **pbWizardJapaneseCalendars**|| **pbWizardJapaneseCatalogs**|| **pbWizardJapaneseCertificates**|| **pbWizardJapaneseEnvelopes**|| **pbWizardJapaneseFlyers**|| **pbWizardJapaneseGiftCertificates**|| **pbWizardJapaneseGreetingCards**|| **pbWizardJapaneseInvitations**|| **pbWizardJapaneseLabels**|| **pbWizardJapaneseLetterheads**|| **pbWizardJapaneseMenus**|| **pbWizardJapaneseNewsletters**|| **pbWizardJapaneseOrigami**|| **pbWizardJapanesePostcards**|| **pbWizardJapanesePrograms**|| **pbWizardJapaneseSigns**|| **pbWizardJapaneseWebSites**|| **pbWizardLabels**|| **pbWizardLetterheads**|| **pbWizardMenus**|| **pbWizardNewsletters**|| **pbWizardNone**|| **pbWizardOrigami**|| **pbWizardPostcards**|| **pbWizardPrograms**|| **pbWizardQuickPublications**|| **pbWizardResumes**|| **pbWizardSigns**|| **pbWizardWebSites**|| **pbWizardWithComplimentsCards**|| **pbWizardWordDocument**|

## <a name="example"></a>Пример

В этом примере создается новая публикация и изменение главной страницы, содержащие номер страницы в звезда в левом верхнем углу страницы.


```vb
Sub CreateNewPublication() 
 Dim AppPub As Application 
 Dim DocPub As Document 
 
 Set AppPub = New Publisher.Application 
 Set DocPub = AppPub.NewDocument 
 AppPub.ActiveWindow.Visible = True 
 
 With DocPub.MasterPages(1).Shapes.AddShape _ 
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
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

