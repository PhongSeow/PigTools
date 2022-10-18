# PigMLang
#### [中文文档](https://github.com/PhongSeow/PigTools/blob/main/Release/PigMLang/README.CN.md)
PigMLang is a very simple multi language solution. It is a method to edit a multi language file with the same name as the application. <br>The file extension is the name or LCID in the following table. <br>It can be stored in the same directory as the application's execution code file. <br>The file content consists of a global "{Global}" and multiple user-defined "{Object Name}". <br>The multi text content consists of multiple "[Key Name]=Text Content"<br>

Example file name:<br>
The application name is PigCmdLib, and the execution code is PigCmdLib.dll, the language region is Chinese (simplified, China), then the multilingual file is PigCmdLib.zh-CN or PigCmdLib.2052.

Example of file content:
```
{Global}
[PressToContinue]=Press any key to continue
[PressYesOrNo]=Press Y to Yes, N to No 
{frmMain}
[Caption]=Form Name
```

Code reference example: 
If the key value cannot find available multilingual text, the content displayed as DefaultText is used.
```
PigMLang.GetMLangText(GlobalKey,DefaultText)
PigMLang.GetMLangText(ObjName,Key,DefaultText)
```

|Name|LCID|EnglishName|DisplayName|NativeName|
| ---- | ---- | ---- | ---- | ---- |
|ar-SA|1025|Arabic (Saudi Arabia)|Arabic (Saudi Arabia)|العربية (المملكة العربية السعودية)|
|bg-BG|1026|Bulgarian (Bulgaria)|Bulgarian (Bulgaria)|български (България)|
|ca-ES|1027|Catalan (Spain)|Catalan (Spain)|català (Espanya)|
|zh-TW|1028|Chinese (Taiwan)|Chinese (Taiwan)|中文（台灣）|
|cs-CZ|1029|Czech (Czech Republic)|Czech (Czech Republic)|čeština (Česká republika)|
|da-DK|1030|Danish (Denmark)|Danish (Denmark)|dansk (Danmark)|
|de-DE|1031|German (Germany)|German (Germany)|Deutsch (Deutschland)|
|el-GR|1032|Greek (Greece)|Greek (Greece)|Ελληνικά (Ελλάδα)|
|en-US|1033|English (United States)|English (United States)|English (United States)|
|es-ES|1034|Spanish (Spain, Sort Order=tradnl)|Spanish (Spain)|Spanish (Spain, Sort Order=tradnl)|
|fi-FI|1035|Finnish (Finland)|Finnish (Finland)|suomi (Suomi)|
|fr-FR|1036|French (France)|French (France)|français (France)|
|he-IL|1037|Hebrew (Israel)|Hebrew (Israel)|עברית (ישראל)|
|hu-HU|1038|Hungarian (Hungary)|Hungarian (Hungary)|magyar (Magyarország)|
|is-IS|1039|Icelandic (Iceland)|Icelandic (Iceland)|íslenska (Ísland)|
|it-IT|1040|Italian (Italy)|Italian (Italy)|italiano (Italia)|
|ja-JP|1041|Japanese (Japan)|Japanese (Japan)|日本語(日本)|
|ko-KR|1042|Korean (South Korea)|Korean (South Korea)|한국어(대한민국)|
|nl-NL|1043|Dutch (Netherlands)|Dutch (Netherlands)|Nederlands (Nederland)|
|nb-NO|1044|Norwegian Bokmål (Norway)|Norwegian Bokmål (Norway)|norsk bokmål (Norge)|
|pl-PL|1045|Polish (Poland)|Polish (Poland)|polski (Polska)|
|pt-BR|1046|Portuguese (Brazil)|Portuguese (Brazil)|português (Brasil)|
|rm-CH|1047|Romansh (Switzerland)|Romansh (Switzerland)|rumantsch (Svizra)|
|ro-RO|1048|Romanian (Romania)|Romanian (Romania)|română (România)|
|ru-RU|1049|Russian (Russia)|Russian (Russia)|русский (Россия)|
|hr-HR|1050|Croatian (Croatia)|Croatian (Croatia)|hrvatski (Hrvatska)|
|sk-SK|1051|Slovak (Slovakia)|Slovak (Slovakia)|slovenčina (Slovensko)|
|sq-AL|1052|Albanian (Albania)|Albanian (Albania)|shqip (Shqipëria)|
|sv-SE|1053|Swedish (Sweden)|Swedish (Sweden)|svenska (Sverige)|
|th-TH|1054|Thai (Thailand)|Thai (Thailand)|ไทย (ไทย)|
|tr-TR|1055|Turkish (Turkey)|Turkish (Turkey)|Türkçe (Türkiye)|
|ur-PK|1056|Urdu (Pakistan)|Urdu (Pakistan)|اردو (پاکستان)|
|id-ID|1057|Indonesian (Indonesia)|Indonesian (Indonesia)|Bahasa Indonesia (Indonesia)|
|uk-UA|1058|Ukrainian (Ukraine)|Ukrainian (Ukraine)|українська (Україна)|
|be-BY|1059|Belarusian (Belarus)|Belarusian (Belarus)|беларуская (Беларусь)|
|sl-SI|1060|Slovenian (Slovenia)|Slovenian (Slovenia)|slovenščina (Slovenija)|
|et-EE|1061|Estonian (Estonia)|Estonian (Estonia)|eesti (Eesti)|
|lv-LV|1062|Latvian (Latvia)|Latvian (Latvia)|latviešu (Latvija)|
|lt-LT|1063|Lithuanian (Lithuania)|Lithuanian (Lithuania)|lietuvių (Lietuva)|
|tg-Cyrl-TJ|1064|Tajik (Cyrillic, Tajikistan)|Tajik (Cyrillic, Tajikistan)|Tajik (Cyrillic, Tajikistan)|
|fa-IR|1065|Persian (Iran)|Persian (Iran)|فارسی (ایران)|
|vi-VN|1066|Vietnamese (Vietnam)|Vietnamese (Vietnam)|Tiếng Việt (Việt Nam)|
|hy-AM|1067|Armenian (Armenia)|Armenian (Armenia)|Հայերէն (Հայաստանի Հանրապետություն)|
|az-Latn-AZ|1068|Azerbaijani (Latin, Azerbaijan)|Azerbaijani (Latin, Azerbaijan)|azərbaycanca (latın, Azərbaycan)|
|eu-ES|1069|Basque (Spain)|Basque (Spain)|euskara (Espainia)|
|hsb-DE|1070|Upper Sorbian (Germany)|Upper Sorbian (Germany)|Upper Sorbian (Germany)|
|mk-MK|1071|Macedonian (Macedonia)|Macedonian (Macedonia)|македонски (Македонија)|
|st-ZA|1072|Southern Sotho (South Africa)|Southern Sotho (South Africa)|Southern Sotho (South Africa)|
|ts-ZA|1073|Tsonga (South Africa)|Tsonga (South Africa)|Tsonga (South Africa)|
|tn-ZA|1074|Tswana (South Africa)|Tswana (South Africa)|Tswana (South Africa)|
|ve-ZA|1075|Venda (South Africa)|Venda (South Africa)|Venda (South Africa)|
|xh-ZA|1076|Xhosa (South Africa)|Xhosa (South Africa)|Xhosa (South Africa)|
|zu-ZA|1077|Zulu (South Africa)|Zulu (South Africa)|isiZulu (iNingizimu Afrika)|
|af-ZA|1078|Afrikaans (South Africa)|Afrikaans (South Africa)|Afrikaans (Suid-Afrika)|
|ka-GE|1079|Georgian (Georgia)|Georgian (Georgia)|ქართული (საქართველო)|
|fo-FO|1080|Faroese (Faroe Islands)|Faroese (Faroe Islands)|føroyskt (Føroyar)|
|hi-IN|1081|Hindi (India)|Hindi (India)|हिन्दी (भारत)|
|mt-MT|1082|Maltese (Malta)|Maltese (Malta)|Malti (Malta)|
|se-NO|1083|Northern Sami (Norway)|Northern Sami (Norway)|Northern Sami (Norway)|
|yi-001|1085|Yiddish (World)|Yiddish (World)|Yiddish (World)|
|ms-MY|1086|Malay (Malaysia)|Malay (Malaysia)|Bahasa Melayu (Malaysia)|
|kk-KZ|1087|Kazakh (Kazakhstan)|Kazakh (Kazakhstan)|қазақ тілі (Қазақстан)|
|ky-KG|1088|Kirghiz (Kyrgyzstan)|Kirghiz (Kyrgyzstan)|Kirghiz (Kyrgyzstan)|
|sw-KE|1089|Swahili (Kenya)|Swahili (Kenya)|Kiswahili (Kenya)|
|tk-TM|1090|Turkmen (Turkmenistan)|Turkmen (Turkmenistan)|Turkmen (Turkmenistan)|
|uz-Latn-UZ|1091|Uzbek (Latin, Uzbekistan)|Uzbek (Latin, Uzbekistan)|oʻzbekcha (Lotin, Oʻzbekiston)|
|tt-RU|1092|Tatar (Russia)|Tatar (Russia)|Tatar (Russia)|
|bn-IN|1093|Bengali (India)|Bengali (India)|বাংলা (ভারত)|
|pa-IN|1094|Punjabi (India)|Punjabi (India)|ਪੰਜਾਬੀ (ਭਾਰਤ)|
|gu-IN|1095|Gujarati (India)|Gujarati (India)|ગુજરાતી (ભારત)|
|or-IN|1096|Oriya (India)|Oriya (India)|ଓଡ଼ିଆ (ଭାରତ)|
|ta-IN|1097|Tamil (India)|Tamil (India)|தமிழ் (இந்தியா)|
|te-IN|1098|Telugu (India)|Telugu (India)|తెలుగు (భారత దేశం)|
|kn-IN|1099|Kannada (India)|Kannada (India)|ಕನ್ನಡ (ಭಾರತ)|
|ml-IN|1100|Malayalam (India)|Malayalam (India)|മലയാളം (ഇന്ത്യ)|
|as-IN|1101|Assamese (India)|Assamese (India)|অসমীয়া (ভাৰত)|
|mr-IN|1102|Marathi (India)|Marathi (India)|मराठी (भारत)|
|sa-IN|1103|Sanskrit (India)|Sanskrit (India)|Sanskrit (India)|
|mn-MN|1104|Mongolian (Mongolia)|Mongolian (Mongolia)|Mongolian (Mongolia)|
|bo-CN|1105|Tibetan (China)|Tibetan (China)|པོད་སྐད་ (རྒྱ་ནག)|
|cy-GB|1106|Welsh (United Kingdom)|Welsh (United Kingdom)|Cymraeg (Prydain Fawr)|
|km-KH|1107|Khmer (Cambodia)|Khmer (Cambodia)|ខ្មែរ (កម្ពុជា)|
|lo-LA|1108|Lao (Laos)|Lao (Laos)|ລາວ (ສ.ປ.ປ ລາວ)|
|my-MM|1109|Burmese (Myanmar [Burma])|Burmese (Myanmar [Burma])|ဗမာ (မြန်မာ)|
|gl-ES|1110|Galician (Spain)|Galician (Spain)|galego (España)|
|kok-IN|1111|Konkani (India)|Konkani (India)|कोंकणी (भारत)|
|mni-IN|1112|Manipuri (India)|Manipuri (India)|Manipuri (India)|
|sd-Deva-IN|1113|Sindhi (Devanagari, India)|Sindhi (Devanagari, India)|Sindhi (Devanagari, India)|
|syr-SY|1114|Syriac (Syria)|Syriac (Syria)|Syriac (Syria)|
|am-ET|1118|Amharic (Ethiopia)|Amharic (Ethiopia)|አማርኛ (ኢትዮጵያ)|
|ks-Arab|1120|Kashmiri (Arabic)|Kashmiri (Arabic)|کٲشُر (اَربی)|
|ne-NP|1121|Nepali (Nepal)|Nepali (Nepal)|नेपाली (नेपाल)|
|fy-NL|1122|Western Frisian (Netherlands)|Western Frisian (Netherlands)|Western Frisian (Netherlands)|
|fil-PH|1124|Filipino (Philippines)|Filipino (Philippines)|Filipino (Pilipinas)|
|bin-NG|1126|Bini (Nigeria)|Bini (Nigeria)|Bini (Nigeria)|
|ig-NG|1136|Igbo (Nigeria)|Igbo (Nigeria)|Igbo (Nigeria)|
|gn-PY|1140|Guarani (Paraguay)|Guarani (Paraguay)|Guarani (Paraguay)|
|la-001|1142|Latin (World)|Latin (World)|Latin (World)|
|so-SO|1143|Somali (Somalia)|Somali (Somalia)|Soomaali (Soomaaliya)|
|mi-NZ|1153|Maori (New Zealand)|Maori (New Zealand)|Maori (New Zealand)|
|ar-IQ|2049|Arabic (Iraq)|Arabic (Iraq)|العربية (العراق)|
|zh-CN|2052|Chinese (China)|Chinese (China)|中文（中国）|
|de-CH|2055|German (Switzerland)|German (Switzerland)|Deutsch (Schweiz)|
|en-GB|2057|English (United Kingdom)|English (United Kingdom)|English (United Kingdom)|
|es-MX|2058|Spanish (Mexico)|Spanish (Mexico)|español (México)|
|fr-BE|2060|French (Belgium)|French (Belgium)|français (Belgique)|
|it-CH|2064|Italian (Switzerland)|Italian (Switzerland)|italiano (Svizzera)|
|nl-BE|2067|Dutch (Belgium)|Dutch (Belgium)|Nederlands (België)|
|nn-NO|2068|Norwegian Nynorsk (Norway)|Norwegian Nynorsk (Norway)|nynorsk (Noreg)|
|pt-PT|2070|Portuguese (Portugal)|Portuguese (Portugal)|português (Portugal)|
|ro-MD|2072|Romanian (Moldova)|Romanian (Moldova)|română (Republica Moldova)|
|ru-MD|2073|Russian (Moldova)|Russian (Moldova)|русский (Молдова)|
|sr-Latn-CS|2074|Serbian (Latin, Serbia)|Serbian (Latin, Serbia)|Srpski (Latinica, Srbija)|
|sv-FI|2077|Swedish (Finland)|Swedish (Finland)|svenska (Finland)|
|az-Cyrl-AZ|2092|Azerbaijani (Cyrillic, Azerbaijan)|Azerbaijani (Cyrillic, Azerbaijan)|Азәрбајҹан (kiril, Азәрбајҹан)|
|ga-IE|2108|Irish (Ireland)|Irish (Ireland)|Gaeilge (Éire)|
|ms-BN|2110|Malay (Brunei)|Malay (Brunei)|Bahasa Melayu (Brunei)|
|uz-Cyrl-UZ|2115|Uzbek (Cyrillic, Uzbekistan)|Uzbek (Cyrillic, Uzbekistan)|Uzbek (Cyrillic, Uzbekistan)|
|bn-BD|2117|Bengali (Bangladesh)|Bengali (Bangladesh)|বাংলা (বাংলাদেশ)|
|mn-Mong-CN|2128|Mongolian (Mongolian, China)|Mongolian (Mongolian, China)|Mongolian (Mongolian, China)|
|ar-EG|3073|Arabic (Egypt)|Arabic (Egypt)|العربية (مصر)|
|zh-HK|3076|Chinese (Hong Kong SAR China)|Chinese (Hong Kong SAR China)|中文（中華人民共和國香港特別行政區）|
|de-AT|3079|German (Austria)|German (Austria)|Deutsch (Österreich)|
|en-AU|3081|English (Australia)|English (Australia)|English (Australia)|
|fr-CA|3084|French (Canada)|French (Canada)|français (Canada)|
|sr-Cyrl-CS|3098|Serbian (Cyrillic, Serbia)|Serbian (Cyrillic, Serbia)|Српски (Ћирилица, Србија)|
|ar-LY|4097|Arabic (Libya)|Arabic (Libya)|العربية (ليبيا)|
|zh-SG|4100|Chinese (Singapore)|Chinese (Singapore)|中文（新加坡）|
|de-LU|4103|German (Luxembourg)|German (Luxembourg)|Deutsch (Luxemburg)|
|en-CA|4105|English (Canada)|English (Canada)|English (Canada)|
|es-GT|4106|Spanish (Guatemala)|Spanish (Guatemala)|español (Guatemala)|
|fr-CH|4108|French (Switzerland)|French (Switzerland)|français (Suisse)|
|ar-DZ|5121|Arabic (Algeria)|Arabic (Algeria)|العربية (الجزائر)|
|zh-MO|5124|Chinese (Macau SAR China)|Chinese (Macau SAR China)|中文（中華人民共和國澳門特別行政區）|
|de-LI|5127|German (Liechtenstein)|German (Liechtenstein)|Deutsch (Liechtenstein)|
|en-NZ|5129|English (New Zealand)|English (New Zealand)|English (New Zealand)|
|es-CR|5130|Spanish (Costa Rica)|Spanish (Costa Rica)|español (Costa Rica)|
|fr-LU|5132|French (Luxembourg)|French (Luxembourg)|français (Luxembourg)|
|bs-Latn-BA|5146|Bosnian (Latin, Bosnia and Herzegovina)|Bosnian (Latin, Bosnia and Herzegovina)|bosanski (latinica, Bosna i Hercegovina)|
|ar-MA|6145|Arabic (Morocco)|Arabic (Morocco)|العربية (المغرب)|
|en-IE|6153|English (Ireland)|English (Ireland)|English (Ireland)|
|es-PA|6154|Spanish (Panama)|Spanish (Panama)|español (Panamá)|
|fr-MC|6156|French (Monaco)|French (Monaco)|français (Monaco)|
|ar-TN|7169|Arabic (Tunisia)|Arabic (Tunisia)|العربية (تونس)|
|en-ZA|7177|English (South Africa)|English (South Africa)|English (South Africa)|
|es-DO|7178|Spanish (Dominican Republic)|Spanish (Dominican Republic)|español (República Dominicana)|
|fr-029|7180|French (Caribbean)|French (Caribbean)|français (Caraïbes)|
|ar-OM|8193|Arabic (Oman)|Arabic (Oman)|العربية (عُمان)|
|en-JM|8201|English (Jamaica)|English (Jamaica)|English (Jamaica)|
|es-VE|8202|Spanish (Venezuela)|Spanish (Venezuela)|español (Venezuela)|
|ar-YE|9217|Arabic (Yemen)|Arabic (Yemen)|العربية (اليمن)|
|en-029|9225|English (Caribbean)|English (Caribbean)|English (Caribbean)|
|es-CO|9226|Spanish (Colombia)|Spanish (Colombia)|español (Colombia)|
|fr-CD|9228|French (Congo - Kinshasa)|French (Congo - Kinshasa)|français (République démocratique du Congo)|
|ar-SY|10241|Arabic (Syria)|Arabic (Syria)|العربية (سوريا)|
|en-BZ|10249|English (Belize)|English (Belize)|English (Belize)|
|es-PE|10250|Spanish (Peru)|Spanish (Peru)|español (Perú)|
|fr-SN|10252|French (Senegal)|French (Senegal)|français (Sénégal)|
|ar-JO|11265|Arabic (Jordan)|Arabic (Jordan)|العربية (الأردن)|
|en-TT|11273|English (Trinidad and Tobago)|English (Trinidad and Tobago)|English (Trinidad and Tobago)|
|es-AR|11274|Spanish (Argentina)|Spanish (Argentina)|español (Argentina)|
|fr-CM|11276|French (Cameroon)|French (Cameroon)|français (Cameroun)|
|ar-LB|12289|Arabic (Lebanon)|Arabic (Lebanon)|العربية (لبنان)|
|en-ZW|12297|English (Zimbabwe)|English (Zimbabwe)|English (Zimbabwe)|
|es-EC|12298|Spanish (Ecuador)|Spanish (Ecuador)|español (Ecuador)|
|fr-CI|12300|French (Côte d’Ivoire)|French (Côte d’Ivoire)|français (Côte d’Ivoire)|
|ar-KW|13313|Arabic (Kuwait)|Arabic (Kuwait)|العربية (الكويت)|
|en-PH|13321|English (Philippines)|English (Philippines)|English (Philippines)|
|es-CL|13322|Spanish (Chile)|Spanish (Chile)|español (Chile)|
|fr-ML|13324|French (Mali)|French (Mali)|français (Mali)|
|ar-AE|14337|Arabic (United Arab Emirates)|Arabic (United Arab Emirates)|العربية (الإمارات العربية المتحدة)|
|es-UY|14346|Spanish (Uruguay)|Spanish (Uruguay)|español (Uruguay)|
|fr-MA|14348|French (Morocco)|French (Morocco)|français (Maroc)|
|ar-BH|15361|Arabic (Bahrain)|Arabic (Bahrain)|العربية (البحرين)|
|es-PY|15370|Spanish (Paraguay)|Spanish (Paraguay)|español (Paraguay)|
|ar-QA|16385|Arabic (Qatar)|Arabic (Qatar)|العربية (قطر)|
|en-IN|16393|English (India)|English (India)|English (India)|
|es-BO|16394|Spanish (Bolivia)|Spanish (Bolivia)|español (Bolivia)|
|es-SV|17418|Spanish (El Salvador)|Spanish (El Salvador)|español (El Salvador)|
|es-HN|18442|Spanish (Honduras)|Spanish (Honduras)|español (Honduras)|
|es-NI|19466|Spanish (Nicaragua)|Spanish (Nicaragua)|español (Nicaragua)|
|es-PR|20490|Spanish (Puerto Rico)|Spanish (Puerto Rico)|español (Puerto Rico)|
