# DocToPlainText

a simple and not so fail proof c# console application, that is able to convert a winword document into a plain text file. 
It is meant to be used in a script that parses a word document.

## Dependencies
`Microsoft.Office.Interop.Word`

The solution depends on an installed word application (it has been tested with winword 2007). The program will start the winword.exe, load the file and invoke the save as command.

## Usage
`DocToPlainText.exe c:/input.doc c:/output.txt`

`DocToPlainText.exe c:/input.doc`

When omitting the second parameter the application will write the contents of the wordfile to the console. This will only work with source files using ASCII characterset.

Specify the output encoding by appending the `enc:xxxx` argument. UTF8 is the default Encoding. 

`DocToPlainText.exe c:/input.doc enc:865`

## Encodings
Here's list of the supported Encodings:

__65001 = UTF8 (default)__    
   37 = EBCDICUSCanada   
  437 = OEMUnitedStates   
  500 = EBCDICInternational   
  708 = ArabicASMO   
  720 = ArabicTransparentASMO   
  737 = OEMGreek437G   
  775 = OEMBaltic   
  850 = OEMMultilingualLatinI   
  852 = OEMMultilingualLatinII   
  855 = OEMCyrillic   
  857 = OEMTurkish   
  860 = OEMPortuguese   
  861 = OEMIcelandic   
  862 = OEMHebrew   
  863 = OEMCanadianFrench   
  864 = OEMArabic   
  865 = OEMNordic   
  866 = OEMCyrillicII   
  869 = OEMModernGreek   
  870 = EBCDICMultilingualROECELatin2   
  874 = Thai   
  875 = EBCDICGreekModern   
  932 = JapaneseShiftJIS   
  936 = SimplifiedChineseGBK   
  949 = Korean   
  950 = TraditionalChineseBig5   
 1026 = EBCDICTurkishLatin5   
 1200 = UnicodeLittleEndian   
 1201 = UnicodeBigEndian   
 1250 = CentralEuropean   
 1251 = Cyrillic   
 1252 = Western   
 1253 = Greek   
 1254 = Turkish   
 1255 = Hebrew   
 1256 = Arabic   
 1257 = Baltic   
 1258 = Vietnamese   
 1361 = KoreanJohab   
10000 = MacRoman   
10001 = MacJapanese   
10002 = MacTraditionalChineseBig5   
10003 = MacKorean   
10004 = MacArabic   
10005 = MacHebrew   
10006 = MacGreek1   
10007 = MacCyrillic   
10008 = MacSimplifiedChineseGB2312   
10010 = MacRomania   
10017 = MacUkraine   
10029 = MacLatin2   
10079 = MacIcelandic   
10081 = MacTurkish   
10082 = MacCroatia   
20000 = TaiwanCNS   
20001 = TaiwanTCA   
20002 = TaiwanEten   
20003 = TaiwanIBM5550   
20004 = TaiwanTeleText   
20005 = TaiwanWang   
20105 = IA5IRV   
20106 = IA5German   
20107 = IA5Swedish   
20108 = IA5Norwegian   
20127 = USASCII   
20261 = T61   
20269 = ISO6937NonSpacingAccent   
20273 = EBCDICGermany   
20277 = EBCDICDenmarkNorway   
20278 = EBCDICFinlandSweden   
20280 = EBCDICItaly   
20284 = EBCDICLatinAmericaSpain   
20285 = EBCDICUnitedKingdom   
20290 = EBCDICJapaneseKatakanaExtended   
20297 = EBCDICFrance   
20420 = EBCDICArabic   
20423 = EBCDICGreek   
20424 = EBCDICHebrew   
20833 = EBCDICKoreanExtended   
20838 = EBCDICThai   
20866 = KOI8R   
20871 = EBCDICIcelandic   
20880 = EBCDICRussian   
20905 = EBCDICTurkish   
21025 = EBCDICSerbianBulgarian   
21027 = ExtAlphaLowercase   
21866 = KOI8U   
28591 = ISO88591Latin1   
28592 = ISO88592CentralEurope   
28593 = ISO88593Latin3   
28594 = ISO88594Baltic   
28595 = ISO88595Cyrillic   
28596 = ISO88596Arabic   
28597 = ISO88597Greek   
28598 = ISO88598Hebrew   
28599 = ISO88599Turkish   
28605 = ISO885915Latin9   
29001 = Europa3   
38598 = ISO88598HebrewLogical   
50001 = AutoDetect   
50220 = ISO2022JPNoHalfwidthKatakana   
50221 = ISO2022JPJISX02021984   
50222 = ISO2022JPJISX02011989   
50225 = ISO2022KR   
50227 = ISO2022CNTraditionalChinese   
50229 = ISO2022CNSimplifiedChinese   
50930 = EBCDICJapaneseKatakanaExtendedAndJapanese   
50931 = EBCDICUSCanadaAndJapanese   
50932 = JapaneseAutoDetect   
50933 = EBCDICKoreanExtendedAndKorean   
50935 = EBCDICSimplifiedChineseExtendedAndSimplifiedChinese   
50936 = SimplifiedChineseAutoDetect   
50937 = EBCDICUSCanadaAndTraditionalChinese   
50939 = EBCDICJapaneseLatinExtendedAndJapanese   
50949 = KoreanAutoDetect   
50950 = TraditionalChineseAutoDetect   
51251 = CyrillicAutoDetect   
51253 = GreekAutoDetect   
51256 = ArabicAutoDetect   
51932 = EUCJapanese   
51936 = EUCChineseSimplifiedChinese   
51949 = EUCKorean   
51950 = EUCTaiwaneseTraditionalChinese   
52936 = HZGBSimplifiedChinese   
54936 = SimplifiedChineseGB18030   
57002 = ISCIIDevanagari   
57003 = ISCIIBengali   
57004 = ISCIITamil   
57005 = ISCIITelugu   
57006 = ISCIIAssamese   
57007 = ISCIIOriya   
57008 = ISCIIKannada   
57009 = ISCIIMalayalam   
57010 = ISCIIGujarati   
57011 = ISCIIPunjabi   
65000 = UTF7   