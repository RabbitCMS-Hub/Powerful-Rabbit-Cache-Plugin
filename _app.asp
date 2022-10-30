<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************

Const	CI_DAKIKA 	= 0
Const	CI_SAAT 	= 1
Const	CI_GUN 		= 2

Class Powerful_Rabbit_Cache_Plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE

	Private ObjSCfso, CacheKomut, fintOnbellekZaman, fintOnbellekAralik, CacheKlasor
	Private fstrSiteAdresi, fstrCacheDosyaUzanti, fstrCekURL, CacheKomutAdi
	Private fstrCacheDosyasi, fstrCachedRamAdi, fstrCachedRamZaman, EsasDosyaAdi, CacheDurumuTrueOrFalse, SETTINGS_SUPER_CACHE_ROOT
	Private exDB, OLUSTURULDU, SUPER_CACHE_FILE_SUBFIX

	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		
		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE&"_")
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "Powerful_Rabbit_Cache_Plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "0")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "0")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", "Powerful-Rabbit-Cache-Plugin")

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Plugin Admin Panel Extention
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "REMOVE:OneFile" Then
			' Call PluginPage("Header")

		    TheFile     = Query.Data("fileName")
		    TheFilePath = SETTINGS_SUPER_CACHE_ROOT & TheFile

		    Set Fso = CreateObject("Scripting.FileSystemObject" )
		    If Fso.FileExists( TheFilePath ) = True then
		        Fso.DeleteFile( TheFilePath ) 
		        Response.Write "{""status"":200}"
		    Else
		        Response.Write "{""status"":404}"
		    End If
		    Set Fso = Nothing 

			' Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "REMOVE:CachedFiles" Then
    		Call PluginPage("Header")
    		
    		Call SuperCacheTemizle()

    		Application.Contents.RemoveAll()

		    Response.Write "<h3>Tüm önbellek dosyaları silindi!</h3>"

		    Call PluginPage("Footer")
		    Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:CachedFiles" Then
    		Call PluginPage("Header")

    		With Response 
				.Write "<link rel=""stylesheet"" href=""/content/plugins/Powerful-Rabbit-Cache-Plugin/style.css"" />"
				.Write "<div class=""form-group"">"
				.Write "    <input type=""text"" class=""form-control"" id=""myInput"" onkeyup=""myFunction()"" placeholder=""Dosya Ara"" />"
				.Write "</div>"
				.Write "<div class=""table-responsive"" style=""max-height:400px;overflow-x:scroll"">"
				.Write "<table id=""myTable"" class=""table table-striped table-bordered table-hovered"">"
				.Write "<tr class=""tableHead"">"
				.Write "    <td>Dosya Adı</td>"
				.Write "    <td>Cache Tarihi</td>"
				.Write "    <td>Güncelleme Tarihi</td>"
				.Write "    <td>Boyut</td>"
				.Write "    <td></td>"
				.Write "</tr>"

				TotalCacheSize = 0
				Set ObjSCfso = CreateObject("Scripting.FileSystemObject")
				Set objKlasor = ObjSCfso.GetFolder(SETTINGS_SUPER_CACHE_ROOT)
				Set colDosyalar = objKlasor.Files

				i=0
				For Each objDosya In colDosyalar
				    TotalCacheSize = TotalCacheSize + objDosya.Size

				.Write "<tr id=""LINE_"& i &""">"
				.Write "    <td><small>"& objDosya.Name &"</small></td>"
				.Write "    <td>"& objDosya.DateCreated &"<br><small>"& NeKadarZamanGecti(objDosya.DateCreated) &"</small></td>"
				.Write "    <td>"& objDosya.DateLastModified &"<br><small>"& NeKadarZamanGecti(objDosya.DateLastModified) &"</small></td>"
				.Write "    <td>"& BoyutHesapla( objDosya.Size ) &"</td>"
				.Write "    <td align=""right"">"
				.Write "    	<div class=""btn-group"">"
				.Write "        	<a class=""btn btn-sm btn-danger btn--icon-text deleteStaticFile"" href=""javascript:void(0)"" data-filename="""& objDosya.Name &""" data-row=""LINE_"& i &""">"
				.Write "            	Sil"
				.Write "        	</a>"
				.Write "    	</div>"
				.Write "    </td>"
				.Write "</tr>"
				    i=i+1
				Next

				Set colDosyalar = nothing
				Set objKlasor = nothing
				Set ObjSCfso = nothing

				.Write "</table>"
				.Write "</div>"
	    		
				.Write "<script>let CACHE_PLUGIN_CODE='"& PLUGIN_CODE &"';</script>"
	    		.Write "<script type=""text/javascript"" src=""/content/plugins/Powerful-Rabbit-Cache-Plugin/js/jquery.min.js""></script>"
	    		.Write "<script type=""text/javascript"" src=""/content/plugins/Powerful-Rabbit-Cache-Plugin/js/app.js""></script>"
    		End With

		    Call PluginPage("Footer")
		    Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "    	<div class=""alert alert-info"">web.config dosyasına yazma izni vermeyi unutmayın</div>"
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("select", ""& PLUGIN_CODE &"_TYPE", "Önbellekleme Türü", "0#Fiziksel Depolama|1#RAM Depolama", TO_FILE)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("select", ""& PLUGIN_CODE &"_TIMING", "Önbellekleme Zamanı", "0#Dakika|1#Saat|2#Gün", TO_FILE)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("number", ""& PLUGIN_CODE &"_TIMING_VAL", "(Dakika,Saat,Gün)' de bir", "", TO_FILE)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("select", ""& PLUGIN_CODE &"_COMPRESSOR", "Belleklenmiş Dosyayı Sıkıştır", "0#Hayır|1#Evet", TO_FILE)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("tag", ""& PLUGIN_CODE &"_EXCLUDE", "Cache Alınmayacak Uzantı ve Adres Belirteçleri", "", TO_FILE)
			.Write "    </div>"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "    	<p class=""alert alert-info cms-style"">rabbitCMS Super Cache sistemi kullanmış olduğunuz CMS için özel olarak geliştirilmiş ön bellekleme modülüdür. Bu modül, web sayfanızın en az güncellenen sayfaların en son halinin statik bir kopyasını anlık olarak adresler ve gelen ziyaretçilere veritabanı bağlantısı dahi sağlanmadan önce sunar. Böylece %95'e kadar daha hızlı bir web sitesi gösterilirken, daha yüksek performanslı çalışan bir sisteme sahip olursunuz. Belleklenen dosyayı sıkıştırma algoritması ile %30'a kadar daha küçük bir çıktı sunarak sunucu-ziyaretçi arasında ki trafiğin hızlanmasını sağlar ve CloudFlare üzerinden servis edilecek HTML/Text türünde sunulur. Sisteme giriş yapmış kullanıcılar ise belirlenen dinamik sayfalar için önbelleklemeyi pas geçerek görüntülerken, yönetici girişi yapanlar her zaman dinamik sonuçları görüntüler.</p>"
			.Write "    </div>"

			.Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:CachedFiles"" class=""btn btn-sm btn-primary"">"
			.Write "        	Önbelleklenmiş Dosyaları Göster"
			.Write "        </a>"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=REMOVE:CachedFiles"" class=""btn btn-sm btn-danger"">"
			.Write "        	Tüm Önbelleği Temizle"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------
	Private Sub Class_Initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_NAME 			= "Powerful Rabbit Cache Plugin"
    	PLUGIN_CODE  			= "SUPER_CACHE"
    	PLUGIN_DB_NAME 			= "plugin_template_db" ' tbl_plugin_XXXXXXX
    	PLUGIN_VERSION 			= "1.3.1"
    	PLUGIN_CREDITS 			= "Coded By @Fatih Aytekin Redevelopment @badursun Anthony Burak DURSUN"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Powerful-Rabbit-Cache-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-http"
    	PLUGIN_REMOVABLE 		= False
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

    	OLUSTURULDU 				= False
    	SETTINGS_SUPER_CACHE_ROOT 	= Server.Mappath( readAppSettings("SUPER_CACHE_PATH") ) & "\"
    	SUPER_CACHE_FILE_SUBFIX 	= readAppSettings("SUPER_CACHE_FILE_SUBFIX")
		
		Set ObjSCfso = CreateObject("Scripting.FileSystemObject")
		
		'Cache Kontrol Değeri
		'------------------------------------------------
		CacheDurumuTrueOrFalse 	= False
		CacheKomutAdi 			= "cache"
		CacheKomut 				= GetQueryString2(CacheKomutAdi)
		'Varsayılan Önbellek Ayarları
		'------------------------------------------------
		fintOnbellekZaman 		= CI_GUN 'GetSettings("SUPER_CACHE_TIMING", "2") 'CI_GUN 
		fintOnbellekAralik 		= 1 'GetSettings("SUPER_CACHE_TIMING_VAL", "1") ' 1
		'Cache Klasörü
		'------------------------------------------------
		CacheKlasor 			= SETTINGS_SUPER_CACHE_ROOT
		fstrSiteAdresi 			= "http://"&Request.ServerVariables("SERVER_NAME")
		If IsHttps = True Then 
			fstrSiteAdresi = "https://"&Request.ServerVariables("SERVER_NAME")
		End If
		'Cache Config
		'------------------------------------------------
		fstrDosyaAdi 			= server.MapPath(Request.ServerVariables("PATH_INFO"))
		fstrTemelAd 			= ObjSCfso.GetBaseName(fstrDosyaAdi) & "_" & HangiSite & "_"
		fstrCacheQueryString 	= GetCacheQueryString()
		fstrCacheDosyaUzanti 	= SUPER_CACHE_FILE_SUBFIX
		'File Cache ' HangiSite
		EsasDosyaAdi 			= Temizle(fstrTemelAd & fstrCacheQueryString & fstrCacheDosyaUzanti,0)
		fstrCacheDosyasi 		= Temizle(CacheKlasor & fstrTemelAd & fstrCacheQueryString & fstrCacheDosyaUzanti,0)
		'RAM Cache
		'------------------------------------------------
		fstrCachedRamAdi 		= Temizle(CStr(fstrTemelAd & fstrCacheQueryString),0)
		fstrCachedRamZaman 		= fstrCachedRamAdi & "_Olustu"
		bIsPostBack 			= false
		If Request("IsPostback") = "true" Then 
			bIsPostBack = true
		End If
		If Request.ServerVariables("HTTP_X-Forwarded-Proto") = "https" Then 
			bIsPostBack = true
		End If

		FullURLParam = request.servervariables("QUERY_STRING")
		If Instr(FullURLParam, ";") <> 0 Then
			fstrCekURL 	= Temizle( Split(request.servervariables("QUERY_STRING"), ";")(1) , 1)
		Else
			fstrCekURL 	= Temizle( fstrSiteAdresi +"/"+ request.servervariables("QUERY_STRING") , 1)
		End If

    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Register App
    	'-------------------------------------------------------------------------------------
    	class_register()
	End Sub
	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------
	Private Sub Class_Terminate()
		If Not ObjSCfso Is Nothing Then
			Set ObjSCfso = Nothing
		End If
	End Sub
	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCode()
		PluginCode = PLUGIN_CODE
	End Property
	Public Property Get PluginName()
		PluginName = PLUGIN_NAME
	End Property
	Public Property Get PluginVersion()
		PluginVersion = PLUGIN_VERSION
	End Property
	Public Property Get PluginGit()
		PluginGit = PLUGIN_GIT
	End Property
	Public Property Get PluginDevURL()
		PluginDevURL = PLUGIN_DEV_URL
	End Property
	Public Property Get PluginFolder()
		PluginFolder = PLUGIN_FILES_ROOT
	End Property
	Public Property Get PluginIcon()
		PluginIcon = PLUGIN_ICON
	End Property
	Public Property Get PluginRemovable()
		PluginRemovable = PLUGIN_REMOVABLE
	End Property
	Public Property Get PluginCredits()
		PluginCredits = PLUGIN_CREDITS
	End Property

	Private Property Get This()
		This = Array(PLUGIN_CODE, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_CREDITS)
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function GetCacheQueryString() 'String Döner
		Dim strSonuc
		Dim var
		Dim strAyir
		strSonuc = vbNullString
		strAyir = vbNullString
		For Each var In Request.QueryString
			If var <> "cache" Then
				strSonuc = strSonuc & strAyir & var & "_" & Request.QueryString(var)
				strAyir = "_"
			End If
		Next
		GetCacheQueryString = strSonuc
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function GetQueryString2(deger)
		If IsNull(deger) Then Exit Function
		strYonURLs = Request.ServerVariables("SERVER_NAME")
		Set strURLs = Request.ServerVariables("QUERY_STRING")

		GetQueryString2 = URLStringParse(strURLs&"&s=x",deger)
		Exit Function
		' Geliştirildiği için burada kesildi...

		arrBolum = Split(strURLs,"/")
		strYonLinks = ""
		For Bo = 3 To UBound(arrBolum)
			If NOT IsEmpty(arrBolum(Bo)) Then 
				Bolum = Split(arrBolum(Bo),"?")
				If UBound(Bolum) > 0 Then
					strIDs = Bolum(0)
				Else
					strIDs = arrBolum(Bo)
				End If
				strYonLinks = strYonLinks & "/" & strIDs
			End If
		Next
		Set strURL = Nothing		
		If IsHttps = True Then 
			If Request.ServerVariables("HTTP_X-Forwarded-Proto") = "https" Then
				' CloudFlare
				strDeger = clearField2(Request.QueryString("404;https://"& strYonURLs &":80"& strYonLinks &"?"& deger &"")) 
			Else
				strDeger = clearField2(Request.QueryString("404;https://"& strYonURLs &":443"& strYonLinks &"?"& deger &"")) 
			End If
		Else 
				strDeger = clearField2(Request.QueryString("404;http://"& strYonURLs &":80"& strYonLinks &"?"& deger &"")) 
		End If

		If strDeger = "" Then 
			GetQueryString2 = clearField2(Request.QueryString(deger)) 
		Else 
			GetQueryString2 = strDeger 
		End If		
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Function ChangeCSRF(HTMLData)
		Dim GelenMetin
			GelenMetin = HTMLData & ""

	    If GelenMetin = "" Then Exit Function
	        
	    Set objRegExp = New Regexp
	    With objRegExp
			.IgnoreCase = True
			.Global 	= True
			.MultiLine 	= True
        	.Pattern 	= "(data-token=')\b([A-Z0-9-]*?)\b(')"
	    End With
	    ChangeCSRF = objRegExp.Replace(HTMLData, "[:CSFRTOKEN:]")
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function ChangeDynamics(HTMLData)
		Dim Data 
			Data = HTMLData & ""

		If IsNull(Data) Then 
			ChangeDynamics = Data
			Exit Function
		End If

		Data = replace(ChangeCSRF(Data), "[:CSFRTOKEN:]", "data-token='"& AntiCSRF.GetCSFRToken() &"'")
		Data = replace(Data, "nocache=", "nocache="& kod(8,"") )
		
		ChangeDynamics = Data
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function IsHttps
		If Request.ServerVariables("HTTP_X-Forwarded-Proto") = "https" Then 
			IsHttps = True
		Else 
			If Request.ServerVariables("HTTPS") = "on" Then IsHttps = True
			If Request.ServerVariables("HTTPS") = "off" Then IsHttps = False
		End If
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function clearField2(str)
		If IsNull(str) Then Exit Function
		str = trim(str)
		str = replace(str,"<","&#60;")
		str = replace(str,"=","&#61;")
		str = replace(str,">","&#62;")
		str = replace(str,"'","&#39;")
		str = replace(str,chr(34),"&#34;")
		str = replace(str,"%","&#37;")
		str = replace(str,"*","&#42;")
		str = replace(str,":","&#58;")
		str = replace(str,"\","&#92;")
		str = replace(str,"/","&#47;")	
		clearField2 = str
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function Temizle(ByVal str, ByVal URL)
		If IsNull(str) Then Exit Function
		
		fstrSiteAdresi = Replace(fstrSiteAdresi, "http://", "https://")
		fstrSiteAdresi = Replace(fstrSiteAdresi, "https://", "http://")
		tmp_p1 = "404_404;"& Trim(fstrSiteAdresi) &":80/"
		tmp_p2 = "404.asp?404;"& fstrSiteAdresi &":80/"
		tmp_p3 = "404_404;"& fstrSiteAdresi &":443/"
		tmp_p4 = "404.asp?404;"& fstrSiteAdresi &":443/"

		str = Replace(str, tmp_p1, "",1,-1,1)
		str = Replace(str, tmp_p2,"",1,-1,1)
		str = Replace(str, tmp_p3, "",1,-1,1)
		str = Replace(str, tmp_p4,"",1,-1,1)
		str = Replace(str, ";https:", "")
		str = Replace(str, ";http:", "")
		str = Replace(str, ":8080", "")
		str = Replace(str, ":80", "")
		str = Replace(str, ":443", "")

		If URL <> 1 Then
			str = Replace(str, "/", "_")
			str = Replace(str, "?", "_")
		End If

		Temizle = str
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function CacheDosyaVarmi()
		CacheDosyaVarmi = ObjSCfso.FileExists(fstrCacheDosyasi)
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function CacheBellektemi()
		If IsEmpty(Application(fstrCachedRamAdi)) OR IsNull(Application(fstrCachedRamAdi)) Then
			CacheBellektemi = false
		Else
			CacheBellektemi = true
		End If
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function CacheDosyaTarihKontrol()
		Dim objDosya
		Dim strTarih
		Dim strBirim
		If CacheDosyaVarmi Then
			Set objDosya = ObjSCfso.GetFile(fstrCacheDosyasi)
				strTarih = objDosya.DateLastModified
			Set objDosya = nothing
			Select Case OnbellekZaman
				Case CI_DAKIKA 	: strBirim = "n"
				Case CI_SAAT 	: strBirim = "h"
				Case CI_GUN 	: strBirim = "d"
				Case Else 		: strBirim = "d"
			End Select
			CacheDosyaTarihKontrol = (DateDiff(strBirim,CDate(strTarih), Now()) > OnbellekAralik)
		Else
			CacheDosyaTarihKontrol = True
		End If
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function CacheRAMTarihKontrol()		
		Dim strBirim
		strBirim = "d"
		If IsEmpty(Application(fstrCachedRamZaman)) OR IsNull(Application(fstrCachedRamZaman)) _
			OR Application(fstrCachedRamZaman) = vbNullString Then
			CacheRAMTarihKontrol = true
		Else
			Select Case OnbellekZaman
				Case CI_DAKIKA
					strBirim = "n"
					
				Case CI_SAAT
					strBirim = "h"
					
				Case CI_GUN
					strBirim = "d"
			End Select
			If CInt(DateDiff(strBirim,CDate(Application(fstrCachedRamZaman)), Now())) > OnbellekAralik Then
				CacheRAMTarihKontrol = true
			Else
				CacheRAMTarihKontrol = false
			End If		
		End If
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Sub DosyadanOku()
		' Cached File Serving Type
		'------------------------------
		If Instr(1, fstrCacheDosyasi, "xml_products.xml") <> 0 Then 
			CACHED_SERVE_TYPE 	= "XML"
			MARKER 				= False
		ElseIf Instr(1, fstrCacheDosyasi, ".json") <> 0 Then 
			CACHED_SERVE_TYPE 	= "JSON"
			MARKER 				= False
		ElseIf Instr(1, fstrCacheDosyasi, ".js") <> 0 Then 
			CACHED_SERVE_TYPE 	= "JS"
			MARKER 				= False
		Else
			CACHED_SERVE_TYPE 	= "HTML"
			MARKER 				= True
		End If
		Select Case CACHED_SERVE_TYPE
			Case "JSON"
				response.ContentType = "application/json"
			Case "JS"
				response.ContentType = "text/javascript"
			Case "XML"
				Response.ContentType = "application/xml"
			Case Else
				Response.ContentType="text/HTML"
		End Select
		
		On Error Resume Next

		Dim objDosya
		Set objDosya = ObjSCfso.OpenTextFile(fstrCacheDosyasi, 1, false, -2)
		
		' CMS İçin NoCache Önleyici
		'Response.Charset 			= "UTF-8"

		' NoCache Replace
		'----------------------------
		Response.Write ChangeDynamics( objDosya.ReadAll )
		' Response.Write Replace( objDosya.ReadAll , "nocache=", "nocache="& kod(8,"") &"")

		If err.number <> 0 Then
            ' Admin Notifications
            '--------------------------------
            AdminNotification "zmdi-explicit", "Rabbit.Cache Error", "Rabbit.Cache plugin bir hata ile karşılaştı ve ön bellek dosyasını sildi.", fstrCacheDosyasi, Response.Status

			OnbellekDosyaSil
			Exit Sub
		End If

		CacheDurumuTrueOrFalse = True
		
		Set objDosya = nothing

		If OLUSTURULDU = True Then 
			Response.Write "<script>let elsCached = document.querySelectorAll('.scripttimer');"
			Response.Write "elsCached.forEach(function(el) {"
			Response.Write "	el.innerHTML='"& FormatNumber(Timer - starttime , 4) &"Sec. w/Cache';"
			Response.Write "});</script>"
		End If
		
		Call SystemTeardown("destroy") ' CMS Destroy
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Sub RAMdanOku()
		CacheDurumuTrueOrFalse = True
		Response.Write Application(fstrCachedRamAdi)
		Call SystemTeardown("destroy")
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Sub Dosya()
		If bIsPostBack Then 
			Exit Sub
		End If

		Select Case CacheKomut		
			Case vbNullString 
				If Not CacheDosyaVarmi() Or CacheDosyaTarihKontrol() Then
					OlusturDosya
				Else
					DosyadanOku
				End If	
			Case "gec" 
			Case "temizle"
				OnbellekDosyaSil			
			Case "htemizle"
				ButunKlasorleriTemizle
			Case "olustur" 
				OlusturDosya
				DosyadanOku					
		End Select
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Function CacheFileName()
		CacheFileName = EsasDosyaAdi
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Function ReadFromCache()
		ReadFromCache = CacheDurumuTrueOrFalse
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Sub Bellek()
		If bIsPostBack Then Exit Sub
		Select Case CacheKomut		
			Case vbNullString 	
				Response.Charset 		= "UTF-8"
				Response.Codepage 		= 65001
				If Not CacheBellektemi() Or CacheRAMTarihKontrol() Then
					RAMeYaz
				Else
					RAMdanOku
				End If
			Case "gec" 							
			Case "temizle"
				Application.Contents.Remove(fstrCachedRamAdi)
			Case "htemizle"
				Application.Contents.RemoveAll()			
			Case "olustur" 
				Response.Charset 		= "UTF-8"
				Response.Codepage 		= 65001
				RAMeYaz
				RAMdanOku				
		End Select
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Sub OlusturDosya()
		OLUSTURULDU = True
		' Type
		If Instr(1, fstrCacheDosyasi, "xml_products.xml") <> 0 Then 
			CACHED_SERVE_TYPE 	= "XML"
			MARKER 				= False
		ElseIf Instr(1, fstrCacheDosyasi, ".json") <> 0 Then 
			CACHED_SERVE_TYPE 	= "JSON"
			MARKER 				= False
		ElseIf Instr(1, fstrCacheDosyasi, ".js") <> 0 Then 
			CACHED_SERVE_TYPE 	= "JS"
			MARKER 				= False
		Else
			CACHED_SERVE_TYPE 	= "HTML"
			MARKER 				= True
		End If

		CACHE_TYPE = Array("Minute", "Hour", "Day")

		cm="<!-- "+vbcrlf
		cm=cm+"-=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=-="+vbcrlf
		cm=cm+",------.         ,--.   ,--.   ,--.  ,--.   ,-----.,--.   ,--. ,---.   "+vbcrlf
		cm=cm+"|  .--. ' ,--,--.|  |-. |  |-. `--',-'  '-.'  .--./|   `.'   |'   .-'  "+vbcrlf
		cm=cm+"|  '--'.'' ,-.  || .-. '| .-. ',--.'-.  .-'|  |    |  |'.'|  |`.  `-.  "+vbcrlf
		cm=cm+"|  |\  \ \ '-'  || `-' || `-' ||  |  |  |  '  '--'\|  |   |  |.-'    | "+vbcrlf
		cm=cm+"`--' '--' `--`--' `---'  `---' `--'  `--'   `-----'`--'   `--'`-----'  "+vbcrlf
		cm=cm+"-=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=-="+vbcrlf
		cm=cm+ "- RabbitCMS SuperCache System Active (c) 2018" + vbcrlf
		cm=cm+ "- This Page Cached @ "& Now() &" "+vbcrlf
		cm=cm+ "- Cache for "& fintOnbellekAralik &" "& CACHE_TYPE(fintOnbellekZaman) &" on pyhisical "+vbcrlf
		If Request.ServerVariables("HTTP_X-Forwarded-Proto") = "https" Then 
		cm=cm+ "- CloudFlare RAYID: "& Request.ServerVariables("HTTP_CF-RAY") &" "+vbcrlf
		End If
		cm=cm+"-=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=-="+vbcrlf
		cm=cm+"-->"+vbcrlf

		Dim strIstek
		Dim objTextStream
		Dim Kaynak
		Dim strCekURL
		strCekURL = fstrCekURL		
		If Instr(1,strCekURL, "?") = 0 Then
			strCekURL = strCekURL & "?"&CacheKomutAdi&"=gec"
		Else
			strCekURL = strCekURL & "&"&CacheKomutAdi&"=gec"
		End If

		strCekURL = Replace(strCekURL, "http://", "https://")

		Set Kaynak = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
			Kaynak.Open "GET", strCekURL, False
			Kaynak.setOption 2, 13056
			Kaynak.setRequestHeader "Content-type", "text/HTML; charset=utf-8"
			Kaynak.setRequestHeader "User-Agent", "Mozilla/4.0+(compatible;+MSIE+7.0;+Windows+NT+5.1)"
			Kaynak.Send()

			'CloudFlare 403 Uyarısı
			'------------------------------------------------
			If Kaynak.Status = 403 Then 
	            ' Admin Notifications
	            '--------------------------------
	            AdminNotification "zmdi-cloud-off", "CloudFlare Banned", "CloudFlare proxy tarafından sunucu IP adresi bloklandığı için veri erişimi yapılamadı. "& Request.ServerVariables("LOCAL_ADDR") &" adresi için Firewall White List eklemesi yapın.", strCekURL, Response.Status

				Response.Write "CloudFlare Banned Server IP.<br />Resolve: Disable Cache Or Add CloudFlare WhiteList server IP.<br>Server IP is: "& Request.ServerVariables("LOCAL_ADDR") &""
				Call SystemTeardown("destroy")
			End If

			If Kaynak.Status = 200 Then
				' const adTypeBinary = 1
				' const adSaveCreateOverwrite = 2
				' const adModeReadWrite = 3

				Set Fso = CreateObject("Scripting.FileSystemObject" )
				If Fso.FileExists( fstrCacheDosyasi ) = True then
					Fso.DeleteFile( fstrCacheDosyasi )
				End If
				Set Fso = Nothing 

				Set objStream = server.CreateObject("ADODB.Stream")
					objStream.Open
					objStream.CharSet = "UTF-8"
					If SETTINGS_SUPER_CACHE_COMPRESSOR = 0 Then 
						objStream.WriteText IIf(MARKER=True, cm, "") + CacheMarker(Kaynak.ResponseText)
					End If
					If SETTINGS_SUPER_CACHE_COMPRESSOR = 1 Then 
						objStream.WriteText IIf(MARKER=True, cm, "") + HTMLtoOneLineCache(CacheMarker(Kaynak.ResponseText))
					End If
					objStream.SaveToFile fstrCacheDosyasi, 2
					objStream.Close
				Set objStream = Nothing

				' Set objTextStream = ObjSCfso.CreateTextFile(fstrCacheDosyasi, True, True)
				' 	If SETTINGS_SUPER_CACHE_COMPRESSOR = 0 Then 
				' 		objTextStream.Write IIf(MARKER=True, cm, "") & CacheMarker(Kaynak.ResponseText)
				' 	End If
				' 	If SETTINGS_SUPER_CACHE_COMPRESSOR = 1 Then 
				' 		objTextStream.Write IIf(MARKER=True, cm, "") & HTMLtoOneLineCache(CacheMarker(Kaynak.ResponseText))
				' 	End If
				' 	objTextStream.Close
				' Set objTextStream = Nothing
			Else
	            ' Admin Notifications
	            '--------------------------------
	            AdminNotification "zmdi-block", "CloudFlare Status Hatası (FILE)", "CloudFlare erişimi denenirken kaynak sunucu "& Kaynak.Status &" status hatası döndü.", strCekURL, Kaynak.Status

				Exit Sub
			End If				
		Set Kaynak = Nothing
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Sub RAMeYaz()
		' Type
		If Instr(1, fstrCacheDosyasi, "xml_products.xml") <> 0 Then 
			CACHED_SERVE_TYPE 	= "XML"
			MARKER 				= False
		ElseIf Instr(1, fstrCacheDosyasi, ".json") <> 0 Then 
			CACHED_SERVE_TYPE 	= "JSON"
			MARKER 				= False
		ElseIf Instr(1, fstrCacheDosyasi, ".js") <> 0 Then 
			CACHED_SERVE_TYPE 	= "JS"
			MARKER 				= False
		Else
			CACHED_SERVE_TYPE 	= "HTML"
			MARKER 				= True
		End If

		CACHE_TYPE = Array("Minute", "Hour", "Day")

		cm="<!-- "+vbcrlf
		cm=cm+"-=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=-="+vbcrlf
		cm=cm+",------.         ,--.   ,--.   ,--.  ,--.   ,-----.,--.   ,--. ,---.   "+vbcrlf
		cm=cm+"|  .--. ' ,--,--.|  |-. |  |-. `--',-'  '-.'  .--./|   `.'   |'   .-'  "+vbcrlf
		cm=cm+"|  '--'.'' ,-.  || .-. '| .-. ',--.'-.  .-'|  |    |  |'.'|  |`.  `-.  "+vbcrlf
		cm=cm+"|  |\  \ \ '-'  || `-' || `-' ||  |  |  |  '  '--'\|  |   |  |.-'    | "+vbcrlf
		cm=cm+"`--' '--' `--`--' `---'  `---' `--'  `--'   `-----'`--'   `--'`-----'  "+vbcrlf
		cm=cm+"-=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=-="+vbcrlf
		cm=cm+ "- RabbitCMS SuperCache System Active (c) 2018" + vbcrlf
		cm=cm+ "- This Page Cached @ "& Now() &" "+vbcrlf
		cm=cm+ "- Cache for "& fintOnbellekAralik &" "& CACHE_TYPE(fintOnbellekZaman) &" on memory "+vbcrlf
		If Request.ServerVariables("HTTP_X-Forwarded-Proto") = "https" Then 
		cm=cm+ "- CloudFlare RAYID: "& Request.ServerVariables("HTTP_CF-RAY") &" "+vbcrlf
		End If
		cm=cm+"-=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=--=*=-="+vbcrlf
		cm=cm+"-->"+vbcrlf

		Dim Kaynak
		Dim strCekURL
		strCekURL = fstrCekURL		
		If Instr(1,strCekURL, "?") = 0 Then
			strCekURL = strCekURL & "?"&CacheKomutAdi&"=gec"
		Else
			strCekURL = strCekURL & "&"&CacheKomutAdi&"=gec"
		End If

		Set Kaynak = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
			Kaynak.Open "GET", strCekURL, False
			Kaynak.setOption 2, 13056
			Kaynak.setRequestHeader "Content-type", "text/HTML; charset=utf-8"
			'Kaynak.setRequestHeader "Content-type", "text/HTML;"
			Kaynak.setRequestHeader "User-Agent", "Mozilla/4.0+(compatible;+MSIE+7.0;+Windows+NT+5.1)"
			Kaynak.Send()

		'CloudFlare 403 Uyarısı
		'------------------------------------------------
		If Kaynak.Status = 403 Then 
            ' Admin Notifications
            '--------------------------------
            If Not TypeName(Func) = "CMSFunctions" Then Set Func = new CMSFunctions
            AdminNotification "zmdi-cloud-off", "CloudFlare Banned", "CloudFlare proxy tarafından sunucu IP adresi bloklandığı için veri erişimi yapılamadı. "& Request.ServerVariables("LOCAL_ADDR") &" adresi için Firewall White List eklemesi yapın.", strCekURL, Response.Status

			Response.Write "CloudFlare Banned Server IP.<br />Resolve: Disable Cache Or Add CloudFlare WhiteList server IP.<br>Server IP is: "& Request.ServerVariables("LOCAL_ADDR") &""
			Call SystemTeardown("destroy")
		End If

		If Kaynak.Status = 200 Then
			If SETTINGS_SUPER_CACHE_COMPRESSOR = 0 Then 
				Application(fstrCachedRamAdi) = IIf(MARKER=True, cm, "") + CacheMarker(Kaynak.ResponseText)
			End If
			If SETTINGS_SUPER_CACHE_COMPRESSOR = 1 Then 
				Application(fstrCachedRamAdi) =  IIf(MARKER=True, cm, "") + HTMLtoOneLineCache(CacheMarker(Kaynak.ResponseText))
			End If
			Application(fstrCachedRamZaman) = Cstr(NOW())
		Else
            ' Admin Notifications
            '--------------------------------
            AdminNotification "zmdi-block", "CloudFlare Status Hatası (RAM)", "CloudFlare erişimi denenirken kaynak sunucu "& Kaynak.Status &" status hatası döndü.", strCekURL, Kaynak.Status

			Exit Sub
		End If
		Set Kaynak = nothing
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Sub OnbellekDosyaSil()
		If CacheDosyaVarmi() Then ObjSCfso.DeleteFile(fstrCacheDosyasi)
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Sub ButunKlasorleriTemizle()
		Dim objKlasor
		Dim colDosyalar
		Dim objDosya
		Set objKlasor = ObjSCfso.GetFolder(SETTINGS_SUPER_CACHE_ROOT)
		Set colDosyalar = objKlasor.Files
			For Each objDosya In colDosyalar
				objDosya.Delete()
			Next
		Set colDosyalar = nothing
		Set objKlasor = nothing
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get OnbellekZaman()
		OnbellekZaman = fintOnbellekZaman
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Let OnbellekZaman(ByVal intYeniDeger)
		fintOnbellekZaman = CInt(intYeniDeger)
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get OnbellekAralik()
		OnbellekAralik = fintOnbellekAralik
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Let OnbellekAralik(ByVal intYeniDeger)
		fintOnbellekAralik = CInt(intYeniDeger)
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get DosyaAdi()
		OnBellekDosyaAdi = fstrCacheDosyasi
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Let OnBellekDosyaAdi(ByVal strYeniDeger)
		fstrCacheDosyasi = CStr(strYeniDeger)
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Function HTMLtoOneLineCache(veri)
	    veri = veri & ""
	    veri = Replace(veri, vbcrlf, "",1,-1,1)
	    veri = Replace(veri, vbcr, "",1,-1,1)
	    veri = Replace(veri, vblf, "",1,-1,1)
	    veri = Replace(veri, vbTab, "",1,-1,1)
	    veri = Replace(veri, "         ", "",1,-1,1)
	    veri = Replace(veri, "        ", "",1,-1,1)
	    veri = Replace(veri, "      ", "",1,-1,1)
	    veri = Replace(veri, "    ", "",1,-1,1)    
	    veri = Replace(veri, "   ", "",1,-1,1)    
	    veri = Replace(veri, "\n", "",1,-1,1)
	    veri = Replace(veri, "\r", "",1,-1,1)
	    veri = Replace(veri, """: """, """:""",1,-1,1)
	    veri = Replace(veri, """ />", """>",1,-1,1)
	    HTMLtoOneLineCache = HTMLCommentTemizleCache( veri )
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Function CacheMarker(veri)
	    veri = veri & ""
	    veri = Replace(veri, "<html ", "<html data-https="""& IsHttps &""" data-supercache=""true"" ",1,-1,1)
	    veri = Replace(veri, "<rss ", "<rss data-supercache=""true"" ",1,-1,1)
	    veri = Replace(veri, "{""LANG_ID"": ", "{""cache_status"":true, ""LANG_ID"": ",1,-1,1)
	    veri = Replace(veri, "xCms;", "xCms; let supercache=true;",1,-1,1)
	    If IsHttps = True Then 
	    	veri = Replace(veri, "http://", "https://",1,-1,1)
	    End If
	    ' CacheMarker = TurkceKarakterEmail( veri )
	    CacheMarker = veri
	End Function

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Function HTMLCommentTemizleCache(GelenMetin)
	    If GelenMetin = "" Then Exit Function
	    
	    Set objRegExp = New Regexp
	    With objRegExp
			.Pattern = "<!--(?!<!)[^\[>].*?-->"
			.IgnoreCase = False
			.Global = True
	    End With
	    
	    HTMLCommentTemizleCache = objRegExp.Replace(GelenMetin,"")
	    
	    Set objRegExp = Nothing
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
End Class
%>
