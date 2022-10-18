# PigMLang
#### [中文文档](https://github.com/PhongSeow/PigTools/blob/master/README.CN.md)
PigMLang是一个非常简洁的多言语解决方案，方法编辑一个与应用同名的多语言文件，文件扩展名为下表中的名称或LCID，与应用的执行码文件存放在相同目录即可，文件内容由一个全局的“{Global}”和多个自定义的“{对象名}”组成，多文本内容由多个“[键名]=文本内容”组成

文件名示例：
应用名称为 PigCmdLib，执行码为 PigCmdLib.dll，语言区域为中文(简体，中国) ，则多语言文件为 PigCmdLib.zh-CN 或 PigCmdLib.2052。

文件内容示例：
{Global}
[PressToContinue]=Press any key to continue
[PressYesOrNo]=Press Y to Yes, N to No
{frmMain}
[Caption]=Form Name 

代码引用示例：
PigMLang.GetMLangText(GlobalKey,DefaultText)
PigMLang.GetMLangText(ObjName,Key,DefaultText)

|名称|LCID|英文名称|显示名称|本国名称|
| ---- | ---- | ---- | ---- | ---- |
|ar-SA|1025|Arabic (Saudi Arabia)|阿拉伯语(沙特阿拉伯)|العربية (المملكة العربية السعودية)|
|bg-BG|1026|Bulgarian (Bulgaria)|保加利亚语(保加利亚)|български (България)|
|ca-ES|1027|Catalan (Catalan)|加泰罗尼亚语 (加泰罗尼亚语)|català (català)|
|zh-TW|1028|Chinese (Traditional, Taiwan)|中文(繁体，中国台湾)|中文(台灣)|
|cs-CZ|1029|Czech (Czechia)|捷克语(捷克共和国)|čeština (Česká republika)|
|da-DK|1030|Danish (Denmark)|丹麦语(丹麦)|dansk (Danmark)|
|de-DE|1031|German (Germany)|德语(德国)|Deutsch (Deutschland)|
|el-GR|1032|Greek (Greece)|希腊语(希腊)|Ελληνικά (Ελλάδα)|
|en-US|1033|English (United States)|英语(美国)|English (United States)|
|es-ES|1034|Spanish (Spain, Traditional Sort)|西班牙语(西班牙)|español (España, alfabetización tradicional)|
|fi-FI|1035|Finnish (Finland)|芬兰语(芬兰)|suomi (Suomi)|
|fr-FR|1036|French (France)|法语(法国)|français (France)|
|he-IL|1037|Hebrew (Israel)|希伯来语(以色列)|עברית (ישראל)|
|hu-HU|1038|Hungarian (Hungary)|匈牙利语(匈牙利)|magyar (Magyarország)|
|is-IS|1039|Icelandic (Iceland)|冰岛语(冰岛)|íslenska (ísland)|
|it-IT|1040|Italian (Italy)|意大利语(意大利)|italiano (Italia)|
|ja-JP|1041|Japanese (Japan)|日语(日本)|日本語 (日本)|
|ko-KR|1042|Korean (Korea)|韩语(韩国)|한국어(대한민국)|
|nl-NL|1043|Dutch (Netherlands)|荷兰语(荷兰)|Nederlands (Nederland)|
|nb-NO|1044|Norwegian Bokmål (Norway)|挪威语、博克马尔语(挪威)|norsk bokmål (Norge)|
|pl-PL|1045|Polish (Poland)|波兰语(波兰)|polski (Polska)|
|pt-BR|1046|Portuguese (Brazil)|葡萄牙语(巴西)|português (Brasil)|
|rm-CH|1047|Romansh (Switzerland)|罗曼什语(瑞士)|rumantsch (Svizra)|
|ro-RO|1048|Romanian (Romania)|罗马尼亚语(罗马尼亚)|română (România)|
|ru-RU|1049|Russian (Russia)|俄语(俄罗斯)|русский (Россия)|
|hr-HR|1050|Croatian (Croatia)|克罗地亚语(克罗地亚)|hrvatski (Hrvatska)|
|sk-SK|1051|Slovak (Slovakia)|斯洛伐克语(斯洛伐克)|slovenčina (Slovensko)|
|sq-AL|1052|Albanian (Albania)|阿尔巴尼亚语(阿尔巴尼亚)|shqip (Shqipëria)|
|sv-SE|1053|Swedish (Sweden)|瑞典语(瑞典)|svenska (Sverige)|
|th-TH|1054|Thai (Thailand)|泰语(泰国)|ไทย (ไทย)|
|tr-TR|1055|Turkish (Turkey)|土耳其语(土耳其)|Türkçe (Türkiye)|
|ur-PK|1056|Urdu (Pakistan)|乌尔都语(巴基斯坦伊斯兰共和国)|اردو (پاکستان)|
|id-ID|1057|Indonesian (Indonesia)|印度尼西亚语(印度尼西亚)|Indonesia (Indonesia)|
|uk-UA|1058|Ukrainian (Ukraine)|乌克兰语(乌克兰)|українська (Україна)|
|be-BY|1059|Belarusian (Belarus)|白俄罗斯语(白俄罗斯)|беларуская (Беларусь)|
|sl-SI|1060|Slovenian (Slovenia)|斯洛文尼亚语(斯洛文尼亚)|slovenščina (Slovenija)|
|et-EE|1061|Estonian (Estonia)|爱沙尼亚语(爱沙尼亚)|eesti (Eesti)|
|lv-LV|1062|Latvian (Latvia)|拉脱维亚语(拉脱维亚)|latviešu (Latvija)|
|lt-LT|1063|Lithuanian (Lithuania)|立陶宛语(立陶宛)|lietuvių (Lietuva)|
|tg-Cyrl-TJ|1064|Tajik (Cyrillic, Tajikistan)|塔吉克语(西里尔文，塔吉克斯坦)|Tajik (Cyrillic, Tajikistan)|
|fa-IR|1065|Persian (Iran)|波斯语(伊朗)|فارسی (ایران)|
|vi-VN|1066|Vietnamese (Vietnam)|越南语(越南)|Tiếng Việt (Việt Nam)|
|hy-AM|1067|Armenian (Armenia)|亚美尼亚语(亚美尼亚)|Հայերէն (Հայաստանի Հանրապետություն)|
|az-Latn-AZ|1068|Azerbaijani (Latin, Azerbaijan)|阿塞拜疆语（拉丁语，阿塞拜疆）|azərbaycanca (latın, Azərbaycan)|
|eu-ES|1069|Basque (Basque)|巴斯克语(巴斯克语)|euskara (euskara)|
|hsb-DE|1070|Upper Sorbian (Germany)|上索布语(德国)|Upper Sorbian (Germany)|
|mk-MK|1071|Macedonian (North Macedonia)|马其顿语(北马其顿)|македонски (Македонија)|
|st-ZA|1072|Sesotho (South Africa)|南索托语(南非)|Sesotho (South Africa)|
|ts-ZA|1073|Xitsonga (South Africa)|汤加语(南非)|Xitsonga (South Africa)|
|tn-ZA|1074|Setswana (South Africa)|茨瓦纳语(南非)|Setswana (Aforika Borwa)|
|ve-ZA|1075|Venda (South Africa)|文达语(南非)|Venda (South Africa)|
|xh-ZA|1076|isiXhosa (South Africa)|科萨语(南非)|isiXhosa (eMzantsi Afrika)|
|zu-ZA|1077|isiZulu (South Africa)|祖鲁语(南非)|isiZulu (iNingizimu Afrika)|
|af-ZA|1078|Afrikaans (South Africa)|南非荷兰语(南非)|Afrikaans (Suid-Afrika)|
|ka-GE|1079|Georgian (Georgia)|格鲁吉亚语(格鲁吉亚)|ქართული (საქართველო)|
|fo-FO|1080|Faroese (Faroe Islands)|法罗语(法罗群岛)|føroyskt (Føroyar)|
|hi-IN|1081|Hindi (India)|印地语(印度)|हिन्दी (भारत)|
|mt-MT|1082|Maltese (Malta)|马耳他语(马耳他)|Malti (Malta)|
|se-NO|1083|Sami, Northern (Norway)|北萨米语(挪威)|davvisámegiella (Norga)|
|yi-001|1085|Yiddish (World)|意第绪语(全球)|Yiddish (World)|
|ms-MY|1086|Malay (Malaysia)|马来语(马来西亚)|Melayu (Malaysia)|
|kk-KZ|1087|Kazakh (Kazakhstan)|哈萨克语(哈萨克斯坦)|қазақ тілі (Қазақстан)|
|ky-KG|1088|Kyrgyz (Kyrgyzstan)|柯尔克孜语(吉尔吉斯坦)|кыргызча (Кыргызстан)|
|sw-KE|1089|Kiswahili (Kenya)|斯瓦希里语(肯尼亚)|Kiswahili (Kenya)|
|tk-TM|1090|Turkmen (Turkmenistan)|土库曼语(土库曼斯坦)|türkmen dili (Türkmenistan)|
|uz-Latn-UZ|1091|Uzbek (Latin, Uzbekistan)|乌兹别克语(拉丁语、乌兹别克斯坦)|oʻzbekcha (Lotin, Oʻzbekiston)|
|tt-RU|1092|Tatar (Russia)|鞑靼语(俄罗斯)|татар (Россия)|
|bn-IN|1093|Bengali (India)|孟加拉语（印度）|বাংলা (ভারত)|
|pa-IN|1094|Punjabi (India)|旁遮普语(印度)|ਪੰਜਾਬੀ (ਭਾਰਤ)|
|gu-IN|1095|Gujarati (India)|古吉拉特语(印度)|ગુજરાતી (ભારત)|
|or-IN|1096|Odia (India)|奥里亚语（印度）|ଓଡ଼ିଆ (ଭାରତ)|
|ta-IN|1097|Tamil (India)|泰米尔语(印度)|தமிழ் (இந்தியா)|
|te-IN|1098|Telugu (India)|泰卢固语(印度)|తెలుగు (భారత దేశం)|
|kn-IN|1099|Kannada (India)|卡纳达语(印度)|ಕನ್ನಡ (ಭಾರತ)|
|ml-IN|1100|Malayalam (India)|马拉雅拉姆语(印度)|മലയാളം (ഇന്ത്യ)|
|as-IN|1101|Assamese (India)|阿萨姆语(印度)|অসমীয়া (ভাৰত)|
|mr-IN|1102|Marathi (India)|马拉地语(印度)|मराठी (भारत)|
|sa-IN|1103|Sanskrit (India)|梵语(印度)|Sanskrit (India)|
|mn-MN|1104|Mongolian (Mongolia)|蒙古语(西里尔语、蒙古)|Mongolian (Mongolia)|
|bo-CN|1105|Tibetan (China)|藏语(中国)|པོད་སྐད་ (རྒྱ་ནག)|
|cy-GB|1106|Welsh (United Kingdom)|威尔士语(英国)|Cymraeg (Y Deyrnas Unedig)|
|km-KH|1107|Khmer (Cambodia)|高棉语(柬埔寨)|ខ្មែរ (កម្ពុជា)|
|lo-LA|1108|Lao (Laos)|老挝语(老挝人民民主共和国)|ລາວ (ສ.ປ.ປ ລາວ)|
|my-MM|1109|Burmese (Myanmar)|缅甸语(缅甸)|ဗမာ (မြန်မာ)|
|gl-ES|1110|Galician (Galician)|加利西亚语(加利西亚语)|galego (galego)|
|kok-IN|1111|Konkani (India)|贡根语(印度)|कोंकणी (भारत)|
|mni-IN|1112|Manipuri (Bangla, India)|曼尼普尔语(印度)|Manipuri (India)|
|sd-Deva-IN|1113|Sindhi (Devanagari, India)|信德语(梵文，印度)|Sindhi (Devanagari, India)|
|syr-SY|1114|Syriac (Syria)|叙利亚语(叙利亚)|Syriac (Syria)|
|am-ET|1118|Amharic (Ethiopia)|阿姆哈拉语(埃塞俄比亚)|አማርኛ (ኢትዮጵያ)|
|ks-Arab|1120|Kashmiri (Arabic)|克什米尔语(波斯阿拉伯文)|کٲشُر (اَربی)|
|ne-NP|1121|Nepali (Nepal)|尼泊尔语(尼泊尔)|नेपाली (नेपाल)|
|fy-NL|1122|Western Frisian (Netherlands)|弗里西亚语(荷兰)|Frysk (Nederlan)|
|fil-PH|1124|Filipino (Philippines)|菲律宾语(菲律宾)|Filipino (Pilipinas)|
|bin-NG|1126|Edo (Nigeria)|克瓦语(尼日利亚)|Bini (Nigeria)|
|ig-NG|1136|Igbo (Nigeria)|伊博语(尼日利亚)|Igbo (Nigeria)|
|gn-PY|1140|Guarani (Paraguay)|瓜拉尼语(巴拉圭)|Guarani (Paraguay)|
|la-VA|1142|Latin (Vatican City)|拉丁语(梵蒂冈)|Latina (Civitas Vaticana)|
|so-SO|1143|Somali (Somalia)|索马里语(索马里)|Soomaali (Soomaaliya)|
|mi-NZ|1153|Maori (New Zealand)|毛利语(新西兰)|te reo Māori (Aotearoa)|
|ar-IQ|2049|Arabic (Iraq)|阿拉伯语(伊拉克)|العربية (العراق)|
|zh-CN|2052|Chinese (Simplified, China)|中文(简体，中国)|中文(中国)|
|de-CH|2055|German (Switzerland)|德语(瑞士)|Deutsch (Schweiz)|
|en-GB|2057|English (United Kingdom)|英语(英国)|English (United Kingdom)|
|es-MX|2058|Spanish (Mexico)|西班牙语(墨西哥)|español (México)|
|fr-BE|2060|French (Belgium)|法语(比利时)|français (Belgique)|
|it-CH|2064|Italian (Switzerland)|意大利语(瑞士)|italiano (Svizzera)|
|nl-BE|2067|Dutch (Belgium)|荷兰语(比利时)|Nederlands (België)|
|nn-NO|2068|Norwegian Nynorsk (Norway)|挪威语、尼诺斯克语(挪威)|norsk nynorsk (Noreg)|
|pt-PT|2070|Portuguese (Portugal)|葡萄牙语(葡萄牙)|português (Portugal)|
|ro-MD|2072|Romanian (Moldova)|罗马尼亚语(摩尔多瓦)|română (Republica Moldova)|
|ru-MD|2073|Russian (Moldova)|俄语(摩尔多瓦)|русский (Молдова)|
|sr-Latn-CS|2074|Serbian (Latin, Serbia and Montenegro (Former))|塞尔维亚语(拉丁语，塞尔维亚和黑山(前))|Srpski (Latinica, Srbija)|
|sv-FI|2077|Swedish (Finland)|瑞典语(芬兰)|svenska (Finland)|
|az-Cyrl-AZ|2092|Azerbaijani (Cyrillic, Azerbaijan)|阿塞拜疆语（西里尔文，阿塞拜疆）|Азәрбајҹан (kiril, Азәрбајҹан)|
|ga-IE|2108|Irish (Ireland)|爱尔兰语(爱尔兰)|Gaeilge (éire)|
|ms-BN|2110|Malay (Brunei)|马来语(文莱达鲁萨兰)|Melayu (Brunei)|
|uz-Cyrl-UZ|2115|Uzbek (Cyrillic, Uzbekistan)|乌兹别克语(西里尔语、乌兹别克斯坦)|Uzbek (Cyrillic, Uzbekistan)|
|bn-BD|2117|Bangla (Bangladesh)|孟加拉语（孟加拉国）|বাংলা (বাংলাদেশ)|
|mn-Mong-CN|2128|Mongolian (Traditional Mongolian, China)|蒙古语(传统蒙古语，中国)|Mongolian (Mongolian, China)|
|ar-EG|3073|Arabic (Egypt)|阿拉伯语(埃及)|العربية (مصر)|
|zh-HK|3076|Chinese (Traditional, Hong Kong SAR)|中文(繁体，中国香港特别行政区)|中文(香港特別行政區)|
|de-AT|3079|German (Austria)|德语(奥地利)|Deutsch (Österreich)|
|en-AU|3081|English (Australia)|英语(澳大利亚)|English (Australia)|
|fr-CA|3084|French (Canada)|法语(加拿大)|français (Canada)|
|sr-Cyrl-CS|3098|Serbian (Cyrillic, Serbia and Montenegro (Former))|塞尔维亚语(西里尔文，塞尔维亚和黑山(前))|Српски (Ћирилица, Србија)|
|ar-LY|4097|Arabic (Libya)|阿拉伯语(利比亚)|العربية (ليبيا)|
|zh-SG|4100|Chinese (Simplified, Singapore)|中文(简体，新加坡)|中文(新加坡)|
|de-LU|4103|German (Luxembourg)|德语(卢森堡)|Deutsch (Luxemburg)|
|en-CA|4105|English (Canada)|英语(加拿大)|English (Canada)|
|es-GT|4106|Spanish (Guatemala)|西班牙语(危地马拉)|español (Guatemala)|
|fr-CH|4108|French (Switzerland)|法语(瑞士)|français (Suisse)|
|ar-DZ|5121|Arabic (Algeria)|阿拉伯语(阿尔及利亚)|العربية (الجزائر)|
|zh-MO|5124|Chinese (Traditional, Macao SAR)|中文(繁体，中国澳门特别行政区)|中文(澳門特別行政區)|
|de-LI|5127|German (Liechtenstein)|德语(列支敦士登)|Deutsch (Liechtenstein)|
|en-NZ|5129|English (New Zealand)|英语(新西兰)|English (New Zealand)|
|es-CR|5130|Spanish (Costa Rica)|西班牙语(哥斯达黎加)|español (Costa Rica)|
|fr-LU|5132|French (Luxembourg)|法语(卢森堡)|français (Luxembourg)|
|bs-Latn-BA|5146|Bosnian (Latin, Bosnia & Herzegovina)|波斯尼亚语(拉丁语，波斯尼亚和黑塞哥维那)|bosanski (Bosna i Hercegovina)|
|ar-MA|6145|Arabic (Morocco)|阿拉伯语(摩洛哥)|العربية (المغرب)|
|en-IE|6153|English (Ireland)|英语(爱尔兰)|English (Ireland)|
|es-PA|6154|Spanish (Panama)|西班牙语(巴拿马)|español (Panamá)|
|fr-MC|6156|French (Monaco)|法语(摩纳哥)|français (Monaco)|
|ar-TN|7169|Arabic (Tunisia)|阿拉伯语(突尼斯)|العربية (تونس)|
|en-ZA|7177|English (South Africa)|英语(南非)|English (South Africa)|
|es-DO|7178|Spanish (Dominican Republic)|西班牙语(多米尼加共和国)|español (República Dominicana)|
|fr-029|7180|French (Caribbean)|法语(加勒比海)|français (Caraïbes)|
|ar-OM|8193|Arabic (Oman)|阿拉伯语(阿曼)|العربية (عُمان)|
|en-JM|8201|English (Jamaica)|英语(牙买加)|English (Jamaica)|
|es-VE|8202|Spanish (Venezuela)|西班牙语(委内瑞拉玻利瓦尔共和国)|español (Venezuela)|
|ar-YE|9217|Arabic (Yemen)|阿拉伯语(也门)|العربية (اليمن)|
|en-029|9225|English (Caribbean)|英语(加勒比海)|English (Caribbean)|
|es-CO|9226|Spanish (Colombia)|西班牙语(哥伦比亚)|español (Colombia)|
|fr-CD|9228|French Congo (DRC)|法语(刚果民主共和国)|français (Congo, République démocratique du)|
|ar-SY|10241|Arabic (Syria)|阿拉伯语(叙利亚)|العربية (سوريا)|
|en-BZ|10249|English (Belize)|英语(伯利兹)|English (Belize)|
|es-PE|10250|Spanish (Peru)|西班牙语(秘鲁)|español (Perú)|
|fr-SN|10252|French (Senegal)|法语(塞内加尔)|français (Sénégal)|
|ar-JO|11265|Arabic (Jordan)|阿拉伯语(约旦)|العربية (الأردن)|
|en-TT|11273|English (Trinidad & Tobago)|英语(特立尼达和多巴哥)|English (Trinidad & Tobago)|
|es-AR|11274|Spanish (Argentina)|西班牙语(阿根廷)|español (Argentina)|
|fr-CM|11276|French (Cameroon)|法语(喀麦隆)|français (Cameroun)|
|ar-LB|12289|Arabic (Lebanon)|阿拉伯语(黎巴嫩)|العربية (لبنان)|
|en-ZW|12297|English (Zimbabwe)|英语(津巴布韦)|English (Zimbabwe)|
|es-EC|12298|Spanish (Ecuador)|西班牙语(厄瓜多尔)|español (Ecuador)|
|fr-CI|12300|French (Côte d’Ivoire)|法语(科特迪瓦)|français (Côte d’Ivoire)|
|ar-KW|13313|Arabic (Kuwait)|阿拉伯语(科威特)|العربية (الكويت)|
|en-PH|13321|English (Philippines)|英语(菲律宾共和国)|English (Philippines)|
|es-CL|13322|Spanish (Chile)|西班牙语(智利)|español (Chile)|
|fr-ML|13324|French (Mali)|法语(马里)|français (Mali)|
|ar-AE|14337|Arabic (United Arab Emirates)|阿拉伯语(阿拉伯联合酋长国)|العربية (الإمارات العربية المتحدة)|
|es-UY|14346|Spanish (Uruguay)|西班牙语(乌拉圭)|español (Uruguay)|
|fr-MA|14348|French (Morocco)|法语(摩洛哥)|français (Maroc)|
|ar-BH|15361|Arabic (Bahrain)|阿拉伯语(巴林)|العربية (البحرين)|
|es-PY|15370|Spanish (Paraguay)|西班牙语(巴拉圭)|español (Paraguay)|
|ar-QA|16385|Arabic (Qatar)|阿拉伯语(卡塔尔)|العربية (قطر)|
|en-IN|16393|English (India)|英语(印度)|English (India)|
|es-BO|16394|Spanish (Bolivia)|西班牙语(玻利维亚)|español (Bolivia)|
|es-SV|17418|Spanish (El Salvador)|西班牙语(萨尔瓦多)|español (El Salvador)|
|es-HN|18442|Spanish (Honduras)|西班牙语(洪都拉斯)|español (Honduras)|
|es-NI|19466|Spanish (Nicaragua)|西班牙语(尼加拉瓜)|español (Nicaragua)|
|es-PR|20490|Spanish (Puerto Rico)|西班牙语(波多黎各)|español (Puerto Rico)|
