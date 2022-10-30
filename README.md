# Powerful-Rabbit-Cache-Plugin
Powerful Rabbit Cache Plugin

# Kullanımı
```asp
<%
Const	CI_DAKIKA   = 0
Const	CI_SAAT     = 1
Const	CI_GUN      = 2

Set objCache = New Cache
    objCache.OnbellekZaman = CI_GUN ' Önbellek Zaman Tipi
    objCache.OnbellekAralik = 1 ' Önbellek zaman aralığı
    objCache.Dosya()  '// Eğer Önbelleği dosya olarak kaydetmek istiyorsanız
    ' objCache.Bellek() '// Eğer Önbelleği sunucu RAM'ine (Application) kaydetmek istiyorsanız (Tavsiye)
    objCache.DosyaAdi = "cached" '// Eğer önbellek ismi belirtmek istiyorsanız (Tavsiye Etmem)
Set objCache = Nothing
%>
```

# Credits
Coded By @Fatih Aytekin
ReDeveloped By @Anthony Burak DURSUN
