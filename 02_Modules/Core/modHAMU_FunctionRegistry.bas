Attribute VB_Name = "modHAMU_FunctionRegistry"
Option Explicit
Option Private Module
Private mRegisteredLanguage As String
Public Sub HAMU_RegisterFunctions()
    If Not HAMU_HasWorkbook() Then Exit Sub
    Dim lang As String
    lang = HAMU_Setting("Language")
    If mRegisteredLanguage = lang Then Exit Sub
    On Error GoTo Failed
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_URLCoz", description:=HAMU_Text("UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("Kodlanmış URL veya metin"))
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_RengeGoreTopla", description:=HAMU_Text("B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası F9 kullanın."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("Toplanacak sayı aralığı"), HAMU_Text("Dolgu rengi örneği olan tek hücre"))
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_RengeGoreSay", description:=HAMU_Text("B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası F9 kullanın."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("Sayılacak hücre aralığı"), HAMU_Text("Dolgu rengi örneği olan tek hücre"))
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_YerelDosyaYolu", description:=HAMU_Text("Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("Kişisel OneDrive URL veya yerel yol"))
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_GizliLink", description:=HAMU_Text("Hücrenin ilk köprü adresini getirir."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("Köprü içeren hücre"))
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_LinkParamDegeri", description:=HAMU_Text("URL içindeki sorgu parametresinin değerini getirir."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("URL içeren hücre"), HAMU_Text("Parametre adı veya adları içeren aralık"))
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_SayidanMetine", description:=HAMU_Text("Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("Türkçe yazıya çevrilecek sayı"))
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!HAMU_TablodanVeriGetir", description:=HAMU_Text("Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir."), Category:="HAMU Tools", ArgumentDescriptions:=Array(HAMU_Text("Aranan değer"), HAMU_Text("Excel tablosunun adı"), HAMU_Text("Aranacak sütun başlığı"), HAMU_Text("Sonuç sütunu başlığı"), HAMU_Text("Eşleşme yoksa dönecek değer"))
    mRegisteredLanguage = lang
    Exit Sub
Failed:
    Debug.Print "HAMU function registration: " & Err.description
End Sub

