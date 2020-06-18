VERSION 5.00
Begin VB.Form frmFuel 
   Caption         =   "Fuel on Board"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmb4 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox cmb3 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cmb2 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fuel on Board in Pounds:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmFuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb1_Click()

    If cmb1 = "35 - 999" Then
    
        cmb2.Clear
        cmb2.AddItem "35 - 99"
        cmb2.AddItem "100 - 199"
        cmb2.AddItem "200 - 299"
        cmb2.AddItem "300 - 399"
        cmb2.AddItem "400 - 499"
        cmb2.AddItem "500 - 599"
        cmb2.AddItem "600 - 699"
        cmb2.AddItem "700 - 799"
        cmb2.AddItem "800 - 899"
        cmb2.AddItem "900 - 999"
        
    ElseIf cmb1 = "1000 - 1499" Then
        
        cmb2.Clear
        cmb2.AddItem "1000 - 1099"
        cmb2.AddItem "1100 - 1199"
        cmb2.AddItem "1200 - 1299"
        cmb2.AddItem "1300 - 1399"
        cmb2.AddItem "1400 - 1499"
    
    ElseIf cmb1 = "1500 - 2224" Then
        
        cmb2.Clear
        cmb2.AddItem "1500 - 1599"
        cmb2.AddItem "1600 - 1699"
        cmb2.AddItem "1700 - 1799"
        cmb2.AddItem "1800 - 1899"
        cmb2.AddItem "1900 - 1999"
        cmb2.AddItem "2000 - 2099"
        cmb2.AddItem "2100 - 2199"
        cmb2.AddItem "2200 - 2224"

    End If
    
End Sub

Private Sub cmb2_Click()

    If cmb2 = "35 - 99" Then
    
        cmb3.Clear
        cmb3.AddItem "35 - 49"
        cmb3.AddItem "50 - 59"
        cmb3.AddItem "60 - 69"
        cmb3.AddItem "70 - 79"
        cmb3.AddItem "80 - 89"
        cmb3.AddItem "90 - 99"
        
    ElseIf cmb2 = "100 - 199" Then
    
        cmb3.Clear
        cmb3.AddItem "100 - 109"
        cmb3.AddItem "110 - 119"
        cmb3.AddItem "120 - 129"
        cmb3.AddItem "130 - 139"
        cmb3.AddItem "140 - 149"
        cmb3.AddItem "150 - 159"
        cmb3.AddItem "160 - 169"
        cmb3.AddItem "170 - 179"
        cmb3.AddItem "180 - 189"
        cmb3.AddItem "190 - 199"
    
    ElseIf cmb2 = "200 - 299" Then
    
        cmb3.Clear
        cmb3.AddItem "200 - 209"
        cmb3.AddItem "210 - 219"
        cmb3.AddItem "220 - 229"
        cmb3.AddItem "230 - 239"
        cmb3.AddItem "240 - 249"
        cmb3.AddItem "250 - 259"
        cmb3.AddItem "260 - 269"
        cmb3.AddItem "270 - 279"
        cmb3.AddItem "280 - 289"
        cmb3.AddItem "290 - 299"
        
    ElseIf cmb2 = "300 - 399" Then
    
        cmb3.Clear
        cmb3.AddItem "300 - 309"
        cmb3.AddItem "310 - 319"
        cmb3.AddItem "320 - 329"
        cmb3.AddItem "330 - 339"
        cmb3.AddItem "340 - 349"
        cmb3.AddItem "350 - 359"
        cmb3.AddItem "360 - 369"
        cmb3.AddItem "370 - 379"
        cmb3.AddItem "380 - 389"
        cmb3.AddItem "390 - 399"
        
    ElseIf cmb2 = "400 - 499" Then
   
        cmb3.Clear
        cmb3.AddItem "400 - 409"
        cmb3.AddItem "410 - 419"
        cmb3.AddItem "420 - 429"
        cmb3.AddItem "430 - 439"
        cmb3.AddItem "440 - 449"
        cmb3.AddItem "450 - 459"
        cmb3.AddItem "460 - 469"
        cmb3.AddItem "470 - 479"
        cmb3.AddItem "480 - 489"
        cmb3.AddItem "490 - 499"
        
    ElseIf cmb2 = "500 - 599" Then
        
        cmb3.Clear
        cmb3.AddItem "500 - 509"
        cmb3.AddItem "510 - 519"
        cmb3.AddItem "520 - 529"
        cmb3.AddItem "530 - 539"
        cmb3.AddItem "540 - 549"
        cmb3.AddItem "550 - 559"
        cmb3.AddItem "560 - 569"
        cmb3.AddItem "570 - 579"
        cmb3.AddItem "580 - 589"
        cmb3.AddItem "590 - 599"
    
    ElseIf cmb2 = "600 - 699" Then
    
        cmb3.Clear
        cmb3.AddItem "600 - 609"
        cmb3.AddItem "610 - 619"
        cmb3.AddItem "620 - 629"
        cmb3.AddItem "630 - 639"
        cmb3.AddItem "640 - 649"
        cmb3.AddItem "650 - 659"
        cmb3.AddItem "660 - 669"
        cmb3.AddItem "670 - 679"
        cmb3.AddItem "680 - 689"
        cmb3.AddItem "690 - 699"
        
    ElseIf cmb2 = "700 - 799" Then
    
        cmb3.Clear
        cmb3.AddItem "700 - 709"
        cmb3.AddItem "710 - 719"
        cmb3.AddItem "720 - 729"
        cmb3.AddItem "730 - 739"
        cmb3.AddItem "740 - 749"
        cmb3.AddItem "750 - 759"
        cmb3.AddItem "760 - 769"
        cmb3.AddItem "770 - 779"
        cmb3.AddItem "780 - 789"
        cmb3.AddItem "790 - 799"
        
    ElseIf cmb2 = "800 - 899" Then
    
        cmb3.Clear
        cmb3.AddItem "800 - 809"
        cmb3.AddItem "810 - 819"
        cmb3.AddItem "820 - 829"
        cmb3.AddItem "830 - 839"
        cmb3.AddItem "840 - 849"
        cmb3.AddItem "850 - 859"
        cmb3.AddItem "860 - 869"
        cmb3.AddItem "870 - 879"
        cmb3.AddItem "880 - 889"
        cmb3.AddItem "890 - 899"
        
    ElseIf cmb2 = "900 - 999" Then
    
        cmb3.Clear
        cmb3.AddItem "900 - 909"
        cmb3.AddItem "910 - 919"
        cmb3.AddItem "920 - 929"
        cmb3.AddItem "930 - 939"
        cmb3.AddItem "940 - 949"
        cmb3.AddItem "950 - 959"
        cmb3.AddItem "960 - 969"
        cmb3.AddItem "970 - 979"
        cmb3.AddItem "980 - 989"
        cmb3.AddItem "990 - 999"
    
    ElseIf cmb2 = "1000 - 1099" Then
    
        cmb3.Clear
        cmb3.AddItem "1000 - 1009"
        cmb3.AddItem "1010 - 1019"
        cmb3.AddItem "1020 - 1029"
        cmb3.AddItem "1030 - 1039"
        cmb3.AddItem "1040 - 1049"
        cmb3.AddItem "1050 - 1059"
        cmb3.AddItem "1060 - 1069"
        cmb3.AddItem "1070 - 1079"
        cmb3.AddItem "1080 - 1089"
        cmb3.AddItem "1090 - 1099"
    
    ElseIf cmb2 = "1100 - 1199" Then
    
        cmb3.Clear
        cmb3.AddItem "1100 - 1109"
        cmb3.AddItem "1110 - 1119"
        cmb3.AddItem "1120 - 1129"
        cmb3.AddItem "1130 - 1139"
        cmb3.AddItem "1140 - 1149"
        cmb3.AddItem "1150 - 1159"
        cmb3.AddItem "1160 - 1169"
        cmb3.AddItem "1170 - 1179"
        cmb3.AddItem "1180 - 1189"
        cmb3.AddItem "1190 - 1199"
    
    ElseIf cmb2 = "1200 - 1299" Then
    
        cmb3.Clear
        cmb3.AddItem "1200 - 1209"
        cmb3.AddItem "1210 - 1219"
        cmb3.AddItem "1220 - 1229"
        cmb3.AddItem "1230 - 1239"
        cmb3.AddItem "1240 - 1249"
        cmb3.AddItem "1250 - 1259"
        cmb3.AddItem "1260 - 1269"
        cmb3.AddItem "1270 - 1279"
        cmb3.AddItem "1280 - 1289"
        cmb3.AddItem "1290 - 1299"
    
    ElseIf cmb2 = "1300 - 1399" Then
    
        cmb3.Clear
        cmb3.AddItem "1300 - 1309"
        cmb3.AddItem "1310 - 1319"
        cmb3.AddItem "1320 - 1329"
        cmb3.AddItem "1330 - 1339"
        cmb3.AddItem "1340 - 1349"
        cmb3.AddItem "1350 - 1359"
        cmb3.AddItem "1360 - 1369"
        cmb3.AddItem "1370 - 1379"
        cmb3.AddItem "1380 - 1389"
        cmb3.AddItem "1390 - 1399"
    
    ElseIf cmb2 = "1400 - 1499" Then
    
        cmb3.Clear
        cmb3.AddItem "1400 - 1409"
        cmb3.AddItem "1410 - 1419"
        cmb3.AddItem "1420 - 1429"
        cmb3.AddItem "1430 - 1439"
        cmb3.AddItem "1440 - 1449"
        cmb3.AddItem "1450 - 1459"
        cmb3.AddItem "1460 - 1469"
        cmb3.AddItem "1470 - 1479"
        cmb3.AddItem "1480 - 1489"
        cmb3.AddItem "1490 - 1499"
        
    ElseIf cmb2 = "1500 - 1599" Then
    
        cmb3.Clear
        cmb3.AddItem "1500 - 1509"
        cmb3.AddItem "1510 - 1519"
        cmb3.AddItem "1520 - 1529"
        cmb3.AddItem "1530 - 1539"
        cmb3.AddItem "1540 - 1549"
        cmb3.AddItem "1550 - 1559"
        cmb3.AddItem "1560 - 1569"
        cmb3.AddItem "1570 - 1579"
        cmb3.AddItem "1580 - 1589"
        cmb3.AddItem "1590 - 1599"
    
    ElseIf cmb2 = "1600 - 1699" Then
    
        cmb3.Clear
        cmb3.AddItem "1600 - 1609"
        cmb3.AddItem "1610 - 1619"
        cmb3.AddItem "1620 - 1629"
        cmb3.AddItem "1630 - 1639"
        cmb3.AddItem "1640 - 1649"
        cmb3.AddItem "1650 - 1659"
        cmb3.AddItem "1660 - 1669"
        cmb3.AddItem "1670 - 1679"
        cmb3.AddItem "1680 - 1689"
        cmb3.AddItem "1690 - 1699"
    
    ElseIf cmb2 = "1700 - 1799" Then
    
        cmb3.Clear
        cmb3.AddItem "1700 - 1709"
        cmb3.AddItem "1710 - 1719"
        cmb3.AddItem "1720 - 1729"
        cmb3.AddItem "1730 - 1739"
        cmb3.AddItem "1740 - 1749"
        cmb3.AddItem "1750 - 1759"
        cmb3.AddItem "1760 - 1769"
        cmb3.AddItem "1770 - 1779"
        cmb3.AddItem "1780 - 1789"
        cmb3.AddItem "1790 - 1799"
    
    ElseIf cmb2 = "1800 - 1899" Then
    
        cmb3.Clear
        cmb3.AddItem "1800 - 1809"
        cmb3.AddItem "1810 - 1819"
        cmb3.AddItem "1820 - 1829"
        cmb3.AddItem "1830 - 1839"
        cmb3.AddItem "1840 - 1849"
        cmb3.AddItem "1850 - 1859"
        cmb3.AddItem "1860 - 1869"
        cmb3.AddItem "1870 - 1879"
        cmb3.AddItem "1880 - 1889"
        cmb3.AddItem "1890 - 1899"
    
    ElseIf cmb2 = "1900 - 1999" Then
    
        cmb3.Clear
        cmb3.AddItem "1900 - 1909"
        cmb3.AddItem "1910 - 1919"
        cmb3.AddItem "1920 - 1929"
        cmb3.AddItem "1930 - 1939"
        cmb3.AddItem "1940 - 1949"
        cmb3.AddItem "1950 - 1959"
        cmb3.AddItem "1960 - 1969"
        cmb3.AddItem "1970 - 1979"
        cmb3.AddItem "1980 - 1989"
        cmb3.AddItem "1990 - 1999"
    
    ElseIf cmb2 = "2000 - 2099" Then
    
        cmb3.Clear
        cmb3.AddItem "2000 - 2009"
        cmb3.AddItem "2010 - 2019"
        cmb3.AddItem "2020 - 2029"
        cmb3.AddItem "2030 - 2039"
        cmb3.AddItem "2040 - 2049"
        cmb3.AddItem "2050 - 2059"
        cmb3.AddItem "2060 - 2069"
        cmb3.AddItem "2070 - 2079"
        cmb3.AddItem "2080 - 2089"
        cmb3.AddItem "2090 - 2099"
    
    ElseIf cmb2 = "2100 - 2199" Then
    
        cmb3.Clear
        cmb3.AddItem "2100 - 2109"
        cmb3.AddItem "2110 - 2119"
        cmb3.AddItem "2120 - 2129"
        cmb3.AddItem "2130 - 2139"
        cmb3.AddItem "2140 - 2149"
        cmb3.AddItem "2150 - 2159"
        cmb3.AddItem "2160 - 2169"
        cmb3.AddItem "2170 - 2179"
        cmb3.AddItem "2180 - 2189"
        cmb3.AddItem "2190 - 2199"
        
    ElseIf cmb2 = "2200 - 2224" Then
        
        cmb3.Clear
        cmb3.AddItem "2200 - 2209"
        cmb3.AddItem "2210 - 2219"
        cmb3.AddItem "2220 - 2224"

    End If

End Sub

Private Sub cmb3_Click()

    If cmb3 = "35 - 49" Then
    
        cmb4.Clear
        cmb4.AddItem "35"
        cmb4.AddItem "36"
        cmb4.AddItem "37"
        cmb4.AddItem "38"
        cmb4.AddItem "39"
        cmb4.AddItem "40"
        cmb4.AddItem "41"
        cmb4.AddItem "42"
        cmb4.AddItem "43"
        cmb4.AddItem "44"
        cmb4.AddItem "45"
        cmb4.AddItem "46"
        cmb4.AddItem "47"
        cmb4.AddItem "48"
        cmb4.AddItem "49"
    
    ElseIf cmb3 = "50 - 59" Then
    
        cmb4.Clear
        cmb4.AddItem "50"
        cmb4.AddItem "51"
        cmb4.AddItem "52"
        cmb4.AddItem "53"
        cmb4.AddItem "54"
        cmb4.AddItem "55"
        cmb4.AddItem "56"
        cmb4.AddItem "57"
        cmb4.AddItem "58"
        cmb4.AddItem "59"
        
    ElseIf cmb3 = "60 - 69" Then
    
        cmb4.Clear
        cmb4.AddItem "60"
        cmb4.AddItem "61"
        cmb4.AddItem "62"
        cmb4.AddItem "63"
        cmb4.AddItem "64"
        cmb4.AddItem "65"
        cmb4.AddItem "66"
        cmb4.AddItem "67"
        cmb4.AddItem "68"
        cmb4.AddItem "69"
    
    ElseIf cmb3 = "70 - 79" Then
    
        cmb4.Clear
        cmb4.AddItem "70"
        cmb4.AddItem "71"
        cmb4.AddItem "72"
        cmb4.AddItem "73"
        cmb4.AddItem "74"
        cmb4.AddItem "75"
        cmb4.AddItem "76"
        cmb4.AddItem "77"
        cmb4.AddItem "78"
        cmb4.AddItem "79"
    
    ElseIf cmb3 = "80 - 89" Then
    
        cmb4.Clear
        cmb4.AddItem "80"
        cmb4.AddItem "81"
        cmb4.AddItem "82"
        cmb4.AddItem "83"
        cmb4.AddItem "84"
        cmb4.AddItem "85"
        cmb4.AddItem "86"
        cmb4.AddItem "87"
        cmb4.AddItem "88"
        cmb4.AddItem "89"
    
    ElseIf cmb3 = "90 - 99" Then
    
        cmb4.Clear
        cmb4.AddItem "90"
        cmb4.AddItem "91"
        cmb4.AddItem "92"
        cmb4.AddItem "93"
        cmb4.AddItem "94"
        cmb4.AddItem "95"
        cmb4.AddItem "96"
        cmb4.AddItem "97"
        cmb4.AddItem "98"
        cmb4.AddItem "99"
    
    ElseIf cmb3 = "100 - 109" Then
    
        cmb4.Clear
        cmb4.AddItem "100"
        cmb4.AddItem "101"
        cmb4.AddItem "102"
        cmb4.AddItem "103"
        cmb4.AddItem "104"
        cmb4.AddItem "105"
        cmb4.AddItem "106"
        cmb4.AddItem "107"
        cmb4.AddItem "108"
        cmb4.AddItem "109"
    
    ElseIf cmb3 = "110 - 119" Then
    
        cmb4.Clear
        cmb4.AddItem "110"
        cmb4.AddItem "111"
        cmb4.AddItem "112"
        cmb4.AddItem "113"
        cmb4.AddItem "114"
        cmb4.AddItem "115"
        cmb4.AddItem "116"
        cmb4.AddItem "117"
        cmb4.AddItem "118"
        cmb4.AddItem "119"
    
    ElseIf cmb3 = "120 - 129" Then
    
        cmb4.Clear
        cmb4.AddItem "120"
        cmb4.AddItem "121"
        cmb4.AddItem "122"
        cmb4.AddItem "123"
        cmb4.AddItem "124"
        cmb4.AddItem "125"
        cmb4.AddItem "126"
        cmb4.AddItem "127"
        cmb4.AddItem "128"
        cmb4.AddItem "129"
    
    ElseIf cmb3 = "130 - 139" Then
    
        cmb4.Clear
        cmb4.AddItem "130"
        cmb4.AddItem "131"
        cmb4.AddItem "132"
        cmb4.AddItem "133"
        cmb4.AddItem "134"
        cmb4.AddItem "135"
        cmb4.AddItem "136"
        cmb4.AddItem "137"
        cmb4.AddItem "138"
        cmb4.AddItem "139"
    
    ElseIf cmb3 = "140 - 149" Then
    
        cmb4.Clear
        cmb4.AddItem "140"
        cmb4.AddItem "141"
        cmb4.AddItem "142"
        cmb4.AddItem "143"
        cmb4.AddItem "144"
        cmb4.AddItem "145"
        cmb4.AddItem "146"
        cmb4.AddItem "147"
        cmb4.AddItem "148"
        cmb4.AddItem "149"
    
    ElseIf cmb3 = "150 - 159" Then
    
        cmb4.Clear
        cmb4.AddItem "150"
        cmb4.AddItem "151"
        cmb4.AddItem "152"
        cmb4.AddItem "153"
        cmb4.AddItem "154"
        cmb4.AddItem "155"
        cmb4.AddItem "156"
        cmb4.AddItem "157"
        cmb4.AddItem "158"
        cmb4.AddItem "159"
    
    ElseIf cmb3 = "160 - 169" Then
    
        cmb4.Clear
        cmb4.AddItem "160"
        cmb4.AddItem "161"
        cmb4.AddItem "162"
        cmb4.AddItem "163"
        cmb4.AddItem "164"
        cmb4.AddItem "165"
        cmb4.AddItem "166"
        cmb4.AddItem "167"
        cmb4.AddItem "168"
        cmb4.AddItem "169"
    
    ElseIf cmb3 = "170 - 179" Then
    
        cmb4.Clear
        cmb4.AddItem "170"
        cmb4.AddItem "171"
        cmb4.AddItem "172"
        cmb4.AddItem "173"
        cmb4.AddItem "174"
        cmb4.AddItem "175"
        cmb4.AddItem "176"
        cmb4.AddItem "177"
        cmb4.AddItem "178"
        cmb4.AddItem "179"
    
    ElseIf cmb3 = "180 - 189" Then
    
        cmb4.Clear
        cmb4.AddItem "180"
        cmb4.AddItem "181"
        cmb4.AddItem "182"
        cmb4.AddItem "183"
        cmb4.AddItem "184"
        cmb4.AddItem "185"
        cmb4.AddItem "186"
        cmb4.AddItem "187"
        cmb4.AddItem "188"
        cmb4.AddItem "189"
    
    ElseIf cmb3 = "190 - 199" Then
    
        cmb4.Clear
        cmb4.AddItem "190"
        cmb4.AddItem "191"
        cmb4.AddItem "192"
        cmb4.AddItem "193"
        cmb4.AddItem "194"
        cmb4.AddItem "195"
        cmb4.AddItem "196"
        cmb4.AddItem "197"
        cmb4.AddItem "198"
        cmb4.AddItem "199"
    
    ElseIf cmb3 = "200 - 209" Then
    
        cmb4.Clear
        cmb4.AddItem "200"
        cmb4.AddItem "201"
        cmb4.AddItem "202"
        cmb4.AddItem "203"
        cmb4.AddItem "204"
        cmb4.AddItem "205"
        cmb4.AddItem "206"
        cmb4.AddItem "207"
        cmb4.AddItem "208"
        cmb4.AddItem "209"
    
    ElseIf cmb3 = "210 - 219" Then
    
        cmb4.Clear
        cmb4.AddItem "210"
        cmb4.AddItem "211"
        cmb4.AddItem "212"
        cmb4.AddItem "213"
        cmb4.AddItem "214"
        cmb4.AddItem "215"
        cmb4.AddItem "216"
        cmb4.AddItem "217"
        cmb4.AddItem "218"
        cmb4.AddItem "219"
    
    ElseIf cmb3 = "220 - 229" Then
    
        cmb4.Clear
        cmb4.AddItem "220"
        cmb4.AddItem "221"
        cmb4.AddItem "222"
        cmb4.AddItem "223"
        cmb4.AddItem "224"
        cmb4.AddItem "225"
        cmb4.AddItem "226"
        cmb4.AddItem "227"
        cmb4.AddItem "228"
        cmb4.AddItem "229"
    
    ElseIf cmb3 = "230 - 239" Then
    
        cmb4.Clear
        cmb4.AddItem "230"
        cmb4.AddItem "231"
        cmb4.AddItem "232"
        cmb4.AddItem "233"
        cmb4.AddItem "234"
        cmb4.AddItem "235"
        cmb4.AddItem "236"
        cmb4.AddItem "237"
        cmb4.AddItem "238"
        cmb4.AddItem "239"
    
    ElseIf cmb3 = "240 - 249" Then
    
        cmb4.Clear
        cmb4.AddItem "240"
        cmb4.AddItem "241"
        cmb4.AddItem "242"
        cmb4.AddItem "243"
        cmb4.AddItem "244"
        cmb4.AddItem "245"
        cmb4.AddItem "246"
        cmb4.AddItem "247"
        cmb4.AddItem "248"
        cmb4.AddItem "249"
    
    ElseIf cmb3 = "250 - 259" Then
    
        cmb4.Clear
        cmb4.AddItem "250"
        cmb4.AddItem "251"
        cmb4.AddItem "252"
        cmb4.AddItem "253"
        cmb4.AddItem "254"
        cmb4.AddItem "255"
        cmb4.AddItem "256"
        cmb4.AddItem "257"
        cmb4.AddItem "258"
        cmb4.AddItem "259"
        
    ElseIf cmb3 = "260 - 269" Then
    
        cmb4.Clear
        cmb4.AddItem "260"
        cmb4.AddItem "261"
        cmb4.AddItem "262"
        cmb4.AddItem "263"
        cmb4.AddItem "264"
        cmb4.AddItem "265"
        cmb4.AddItem "266"
        cmb4.AddItem "267"
        cmb4.AddItem "268"
        cmb4.AddItem "269"
    
    ElseIf cmb3 = "270 - 279" Then
    
        cmb4.Clear
        cmb4.AddItem "270"
        cmb4.AddItem "271"
        cmb4.AddItem "272"
        cmb4.AddItem "273"
        cmb4.AddItem "274"
        cmb4.AddItem "275"
        cmb4.AddItem "276"
        cmb4.AddItem "277"
        cmb4.AddItem "278"
        cmb4.AddItem "279"
    
    ElseIf cmb3 = "280 - 289" Then
    
        cmb4.Clear
        cmb4.AddItem "280"
        cmb4.AddItem "281"
        cmb4.AddItem "282"
        cmb4.AddItem "283"
        cmb4.AddItem "284"
        cmb4.AddItem "285"
        cmb4.AddItem "286"
        cmb4.AddItem "287"
        cmb4.AddItem "288"
        cmb4.AddItem "289"
    
    ElseIf cmb3 = "290 - 299" Then
    
        cmb4.Clear
        cmb4.AddItem "290"
        cmb4.AddItem "291"
        cmb4.AddItem "292"
        cmb4.AddItem "293"
        cmb4.AddItem "294"
        cmb4.AddItem "295"
        cmb4.AddItem "296"
        cmb4.AddItem "297"
        cmb4.AddItem "298"
        cmb4.AddItem "299"
    
    ElseIf cmb3 = "300 - 309" Then
    
        cmb4.Clear
        cmb4.AddItem "300"
        cmb4.AddItem "301"
        cmb4.AddItem "302"
        cmb4.AddItem "303"
        cmb4.AddItem "304"
        cmb4.AddItem "305"
        cmb4.AddItem "306"
        cmb4.AddItem "307"
        cmb4.AddItem "308"
        cmb4.AddItem "309"
    
    ElseIf cmb3 = "310 - 319" Then
    
        cmb4.Clear
        cmb4.AddItem "310"
        cmb4.AddItem "311"
        cmb4.AddItem "312"
        cmb4.AddItem "313"
        cmb4.AddItem "314"
        cmb4.AddItem "315"
        cmb4.AddItem "316"
        cmb4.AddItem "317"
        cmb4.AddItem "318"
        cmb4.AddItem "319"
    
    ElseIf cmb3 = "320 - 329" Then
    
        cmb4.Clear
        cmb4.AddItem "320"
        cmb4.AddItem "321"
        cmb4.AddItem "322"
        cmb4.AddItem "323"
        cmb4.AddItem "324"
        cmb4.AddItem "325"
        cmb4.AddItem "326"
        cmb4.AddItem "327"
        cmb4.AddItem "328"
        cmb4.AddItem "329"
    
    ElseIf cmb3 = "330 - 339" Then
    
        cmb4.Clear
        cmb4.AddItem "330"
        cmb4.AddItem "331"
        cmb4.AddItem "332"
        cmb4.AddItem "333"
        cmb4.AddItem "334"
        cmb4.AddItem "335"
        cmb4.AddItem "336"
        cmb4.AddItem "337"
        cmb4.AddItem "338"
        cmb4.AddItem "339"
    
    ElseIf cmb3 = "340 - 349" Then
    
        cmb4.Clear
        cmb4.AddItem "340"
        cmb4.AddItem "341"
        cmb4.AddItem "342"
        cmb4.AddItem "343"
        cmb4.AddItem "344"
        cmb4.AddItem "345"
        cmb4.AddItem "346"
        cmb4.AddItem "347"
        cmb4.AddItem "348"
        cmb4.AddItem "349"
    
    ElseIf cmb3 = "350 - 359" Then
    
        cmb4.Clear
        cmb4.AddItem "350"
        cmb4.AddItem "351"
        cmb4.AddItem "352"
        cmb4.AddItem "353"
        cmb4.AddItem "354"
        cmb4.AddItem "355"
        cmb4.AddItem "356"
        cmb4.AddItem "357"
        cmb4.AddItem "358"
        cmb4.AddItem "359"
    
    ElseIf cmb3 = "360 - 369" Then
    
        cmb4.Clear
        cmb4.AddItem "360"
        cmb4.AddItem "361"
        cmb4.AddItem "362"
        cmb4.AddItem "363"
        cmb4.AddItem "364"
        cmb4.AddItem "365"
        cmb4.AddItem "366"
        cmb4.AddItem "367"
        cmb4.AddItem "368"
        cmb4.AddItem "369"
    
    ElseIf cmb3 = "370 - 379" Then
    
        cmb4.Clear
        cmb4.AddItem "370"
        cmb4.AddItem "371"
        cmb4.AddItem "372"
        cmb4.AddItem "373"
        cmb4.AddItem "374"
        cmb4.AddItem "375"
        cmb4.AddItem "376"
        cmb4.AddItem "377"
        cmb4.AddItem "378"
        cmb4.AddItem "379"
    
    ElseIf cmb3 = "380 - 389" Then
    
        cmb4.Clear
        cmb4.AddItem "380"
        cmb4.AddItem "381"
        cmb4.AddItem "382"
        cmb4.AddItem "383"
        cmb4.AddItem "384"
        cmb4.AddItem "385"
        cmb4.AddItem "386"
        cmb4.AddItem "387"
        cmb4.AddItem "388"
        cmb4.AddItem "389"
        
    ElseIf cmb3 = "390 - 399" Then
    
        cmb4.Clear
        cmb4.AddItem "390"
        cmb4.AddItem "391"
        cmb4.AddItem "392"
        cmb4.AddItem "393"
        cmb4.AddItem "394"
        cmb4.AddItem "395"
        cmb4.AddItem "396"
        cmb4.AddItem "397"
        cmb4.AddItem "398"
        cmb4.AddItem "399"
    
    ElseIf cmb3 = "400 - 409" Then
    
        cmb4.Clear
        cmb4.AddItem "400"
        cmb4.AddItem "401"
        cmb4.AddItem "402"
        cmb4.AddItem "403"
        cmb4.AddItem "404"
        cmb4.AddItem "405"
        cmb4.AddItem "406"
        cmb4.AddItem "407"
        cmb4.AddItem "408"
        cmb4.AddItem "409"
        
    
    ElseIf cmb3 = "410 - 419" Then
    
        cmb4.Clear
        cmb4.AddItem "410"
        cmb4.AddItem "411"
        cmb4.AddItem "412"
        cmb4.AddItem "413"
        cmb4.AddItem "414"
        cmb4.AddItem "415"
        cmb4.AddItem "416"
        cmb4.AddItem "417"
        cmb4.AddItem "418"
        cmb4.AddItem "419"
    
    ElseIf cmb3 = "420 - 429" Then
    
        cmb4.Clear
        cmb4.AddItem "420"
        cmb4.AddItem "421"
        cmb4.AddItem "422"
        cmb4.AddItem "423"
        cmb4.AddItem "424"
        cmb4.AddItem "425"
        cmb4.AddItem "426"
        cmb4.AddItem "427"
        cmb4.AddItem "428"
        cmb4.AddItem "429"
    
    ElseIf cmb3 = "430 - 439" Then
    
        cmb4.Clear
        cmb4.AddItem "430"
        cmb4.AddItem "431"
        cmb4.AddItem "432"
        cmb4.AddItem "433"
        cmb4.AddItem "434"
        cmb4.AddItem "435"
        cmb4.AddItem "436"
        cmb4.AddItem "437"
        cmb4.AddItem "438"
        cmb4.AddItem "439"
        
    ElseIf cmb3 = "440 - 449" Then
    
        cmb4.Clear
        cmb4.AddItem "440"
        cmb4.AddItem "441"
        cmb4.AddItem "442"
        cmb4.AddItem "443"
        cmb4.AddItem "444"
        cmb4.AddItem "445"
        cmb4.AddItem "446"
        cmb4.AddItem "447"
        cmb4.AddItem "448"
        cmb4.AddItem "449"
    
    ElseIf cmb3 = "450 - 459" Then
    
        cmb4.Clear
        cmb4.AddItem "450"
        cmb4.AddItem "451"
        cmb4.AddItem "452"
        cmb4.AddItem "453"
        cmb4.AddItem "454"
        cmb4.AddItem "455"
        cmb4.AddItem "456"
        cmb4.AddItem "457"
        cmb4.AddItem "458"
        cmb4.AddItem "459"
    
    ElseIf cmb3 = "460 - 469" Then
    
        cmb4.Clear
        cmb4.AddItem "460"
        cmb4.AddItem "461"
        cmb4.AddItem "462"
        cmb4.AddItem "463"
        cmb4.AddItem "464"
        cmb4.AddItem "465"
        cmb4.AddItem "466"
        cmb4.AddItem "467"
        cmb4.AddItem "468"
        cmb4.AddItem "469"
    
    ElseIf cmb3 = "470 - 479" Then
    
        cmb4.Clear
        cmb4.AddItem "470"
        cmb4.AddItem "471"
        cmb4.AddItem "472"
        cmb4.AddItem "473"
        cmb4.AddItem "474"
        cmb4.AddItem "475"
        cmb4.AddItem "476"
        cmb4.AddItem "477"
        cmb4.AddItem "478"
        cmb4.AddItem "479"
    
    ElseIf cmb3 = "480 - 489" Then
    
        cmb4.Clear
        cmb4.AddItem "480"
        cmb4.AddItem "481"
        cmb4.AddItem "482"
        cmb4.AddItem "483"
        cmb4.AddItem "484"
        cmb4.AddItem "485"
        cmb4.AddItem "486"
        cmb4.AddItem "487"
        cmb4.AddItem "488"
        cmb4.AddItem "489"
    
    ElseIf cmb3 = "490 - 499" Then
    
        cmb4.Clear
        cmb4.AddItem "490"
        cmb4.AddItem "491"
        cmb4.AddItem "492"
        cmb4.AddItem "493"
        cmb4.AddItem "494"
        cmb4.AddItem "495"
        cmb4.AddItem "496"
        cmb4.AddItem "497"
        cmb4.AddItem "498"
        cmb4.AddItem "499"
    
    ElseIf cmb3 = "500 - 509" Then
    
        cmb4.Clear
        cmb4.AddItem "500"
        cmb4.AddItem "501"
        cmb4.AddItem "502"
        cmb4.AddItem "503"
        cmb4.AddItem "504"
        cmb4.AddItem "505"
        cmb4.AddItem "506"
        cmb4.AddItem "507"
        cmb4.AddItem "508"
        cmb4.AddItem "509"
    
    ElseIf cmb3 = "510 - 519" Then
    
        cmb4.Clear
        cmb4.AddItem "510"
        cmb4.AddItem "511"
        cmb4.AddItem "512"
        cmb4.AddItem "513"
        cmb4.AddItem "514"
        cmb4.AddItem "515"
        cmb4.AddItem "516"
        cmb4.AddItem "517"
        cmb4.AddItem "518"
        cmb4.AddItem "519"
    
    ElseIf cmb3 = "520 - 529" Then
    
        cmb4.Clear
        cmb4.AddItem "520"
        cmb4.AddItem "521"
        cmb4.AddItem "522"
        cmb4.AddItem "523"
        cmb4.AddItem "524"
        cmb4.AddItem "525"
        cmb4.AddItem "526"
        cmb4.AddItem "527"
        cmb4.AddItem "528"
        cmb4.AddItem "529"
    
    ElseIf cmb3 = "530 - 539" Then
    
        cmb4.Clear
        cmb4.AddItem "530"
        cmb4.AddItem "531"
        cmb4.AddItem "532"
        cmb4.AddItem "533"
        cmb4.AddItem "534"
        cmb4.AddItem "535"
        cmb4.AddItem "536"
        cmb4.AddItem "537"
        cmb4.AddItem "538"
        cmb4.AddItem "539"
    
    ElseIf cmb3 = "540 - 549" Then
    
        cmb4.Clear
        cmb4.AddItem "540"
        cmb4.AddItem "541"
        cmb4.AddItem "542"
        cmb4.AddItem "543"
        cmb4.AddItem "544"
        cmb4.AddItem "545"
        cmb4.AddItem "546"
        cmb4.AddItem "547"
        cmb4.AddItem "548"
        cmb4.AddItem "549"
    
    ElseIf cmb3 = "550 - 559" Then
    
        cmb4.Clear
        cmb4.AddItem "550"
        cmb4.AddItem "551"
        cmb4.AddItem "552"
        cmb4.AddItem "553"
        cmb4.AddItem "554"
        cmb4.AddItem "555"
        cmb4.AddItem "556"
        cmb4.AddItem "557"
        cmb4.AddItem "558"
        cmb4.AddItem "559"
    
    ElseIf cmb3 = "560 - 569" Then
    
        cmb4.Clear
        cmb4.AddItem "560"
        cmb4.AddItem "561"
        cmb4.AddItem "562"
        cmb4.AddItem "563"
        cmb4.AddItem "564"
        cmb4.AddItem "565"
        cmb4.AddItem "566"
        cmb4.AddItem "567"
        cmb4.AddItem "568"
        cmb4.AddItem "569"
    
    ElseIf cmb3 = "570 - 579" Then
    
        cmb4.Clear
        cmb4.AddItem "570"
        cmb4.AddItem "571"
        cmb4.AddItem "572"
        cmb4.AddItem "573"
        cmb4.AddItem "574"
        cmb4.AddItem "575"
        cmb4.AddItem "576"
        cmb4.AddItem "577"
        cmb4.AddItem "578"
        cmb4.AddItem "579"
    
    ElseIf cmb3 = "580 - 589" Then
    
        cmb4.Clear
        cmb4.AddItem "580"
        cmb4.AddItem "581"
        cmb4.AddItem "582"
        cmb4.AddItem "583"
        cmb4.AddItem "584"
        cmb4.AddItem "585"
        cmb4.AddItem "586"
        cmb4.AddItem "587"
        cmb4.AddItem "588"
        cmb4.AddItem "589"
    
    ElseIf cmb3 = "590 - 599" Then
    
        cmb4.Clear
        cmb4.AddItem "590"
        cmb4.AddItem "591"
        cmb4.AddItem "592"
        cmb4.AddItem "593"
        cmb4.AddItem "594"
        cmb4.AddItem "595"
        cmb4.AddItem "596"
        cmb4.AddItem "597"
        cmb4.AddItem "598"
        cmb4.AddItem "599"
    
    ElseIf cmb3 = "600 - 609" Then
    
        cmb4.Clear
        cmb4.AddItem "600"
        cmb4.AddItem "601"
        cmb4.AddItem "602"
        cmb4.AddItem "603"
        cmb4.AddItem "604"
        cmb4.AddItem "605"
        cmb4.AddItem "606"
        cmb4.AddItem "607"
        cmb4.AddItem "608"
        cmb4.AddItem "609"
    
    ElseIf cmb3 = "610 - 619" Then
    
        cmb4.Clear
        cmb4.AddItem "610"
        cmb4.AddItem "611"
        cmb4.AddItem "612"
        cmb4.AddItem "613"
        cmb4.AddItem "614"
        cmb4.AddItem "615"
        cmb4.AddItem "616"
        cmb4.AddItem "617"
        cmb4.AddItem "618"
        cmb4.AddItem "619"
    
    ElseIf cmb3 = "620 - 629" Then
    
        cmb4.Clear
        cmb4.AddItem "620"
        cmb4.AddItem "621"
        cmb4.AddItem "622"
        cmb4.AddItem "623"
        cmb4.AddItem "624"
        cmb4.AddItem "625"
        cmb4.AddItem "626"
        cmb4.AddItem "627"
        cmb4.AddItem "628"
        cmb4.AddItem "629"
    
    ElseIf cmb3 = "630 - 639" Then
    
        cmb4.Clear
        cmb4.AddItem "630"
        cmb4.AddItem "631"
        cmb4.AddItem "632"
        cmb4.AddItem "633"
        cmb4.AddItem "634"
        cmb4.AddItem "635"
        cmb4.AddItem "636"
        cmb4.AddItem "637"
        cmb4.AddItem "638"
        cmb4.AddItem "639"
    
    ElseIf cmb3 = "640 - 649" Then
    
        cmb4.Clear
        cmb4.AddItem "640"
        cmb4.AddItem "641"
        cmb4.AddItem "642"
        cmb4.AddItem "643"
        cmb4.AddItem "644"
        cmb4.AddItem "645"
        cmb4.AddItem "646"
        cmb4.AddItem "647"
        cmb4.AddItem "648"
        cmb4.AddItem "649"
    
    ElseIf cmb3 = "650 - 659" Then
    
        cmb4.Clear
        cmb4.AddItem "650"
        cmb4.AddItem "651"
        cmb4.AddItem "652"
        cmb4.AddItem "653"
        cmb4.AddItem "654"
        cmb4.AddItem "655"
        cmb4.AddItem "656"
        cmb4.AddItem "657"
        cmb4.AddItem "658"
        cmb4.AddItem "659"
    
    ElseIf cmb3 = "660 - 669" Then
    
        cmb4.Clear
        cmb4.AddItem "660"
        cmb4.AddItem "661"
        cmb4.AddItem "662"
        cmb4.AddItem "663"
        cmb4.AddItem "664"
        cmb4.AddItem "665"
        cmb4.AddItem "666"
        cmb4.AddItem "667"
        cmb4.AddItem "668"
        cmb4.AddItem "669"
    
    ElseIf cmb3 = "670 - 679" Then
    
        cmb4.Clear
        cmb4.AddItem "670"
        cmb4.AddItem "671"
        cmb4.AddItem "672"
        cmb4.AddItem "673"
        cmb4.AddItem "674"
        cmb4.AddItem "675"
        cmb4.AddItem "676"
        cmb4.AddItem "677"
        cmb4.AddItem "678"
        cmb4.AddItem "679"
    
    ElseIf cmb3 = "680 - 689" Then
    
        cmb4.Clear
        cmb4.AddItem "680"
        cmb4.AddItem "681"
        cmb4.AddItem "682"
        cmb4.AddItem "683"
        cmb4.AddItem "684"
        cmb4.AddItem "685"
        cmb4.AddItem "686"
        cmb4.AddItem "687"
        cmb4.AddItem "688"
        cmb4.AddItem "689"
    
    ElseIf cmb3 = "690 - 699" Then
    
        cmb4.Clear
        cmb4.AddItem "690"
        cmb4.AddItem "691"
        cmb4.AddItem "692"
        cmb4.AddItem "693"
        cmb4.AddItem "694"
        cmb4.AddItem "695"
        cmb4.AddItem "696"
        cmb4.AddItem "697"
        cmb4.AddItem "698"
        cmb4.AddItem "699"
    
    ElseIf cmb3 = "700 - 709" Then
    
        cmb4.Clear
        cmb4.AddItem "700"
        cmb4.AddItem "701"
        cmb4.AddItem "702"
        cmb4.AddItem "703"
        cmb4.AddItem "704"
        cmb4.AddItem "705"
        cmb4.AddItem "706"
        cmb4.AddItem "707"
        cmb4.AddItem "708"
        cmb4.AddItem "709"
    
    ElseIf cmb3 = "710 - 719" Then
    
        cmb4.Clear
        cmb4.AddItem "710"
        cmb4.AddItem "711"
        cmb4.AddItem "712"
        cmb4.AddItem "713"
        cmb4.AddItem "714"
        cmb4.AddItem "715"
        cmb4.AddItem "716"
        cmb4.AddItem "717"
        cmb4.AddItem "718"
        cmb4.AddItem "719"
    
    ElseIf cmb3 = "720 - 729" Then
    
        cmb4.Clear
        cmb4.AddItem "720"
        cmb4.AddItem "721"
        cmb4.AddItem "722"
        cmb4.AddItem "723"
        cmb4.AddItem "724"
        cmb4.AddItem "725"
        cmb4.AddItem "726"
        cmb4.AddItem "727"
        cmb4.AddItem "728"
        cmb4.AddItem "729"
    
    ElseIf cmb3 = "730 - 739" Then
    
        cmb4.Clear
        cmb4.AddItem "730"
        cmb4.AddItem "731"
        cmb4.AddItem "732"
        cmb4.AddItem "733"
        cmb4.AddItem "734"
        cmb4.AddItem "735"
        cmb4.AddItem "736"
        cmb4.AddItem "737"
        cmb4.AddItem "738"
        cmb4.AddItem "739"
    
    ElseIf cmb3 = "740 - 749" Then
    
        cmb4.Clear
        cmb4.AddItem "740"
        cmb4.AddItem "741"
        cmb4.AddItem "742"
        cmb4.AddItem "743"
        cmb4.AddItem "744"
        cmb4.AddItem "745"
        cmb4.AddItem "746"
        cmb4.AddItem "747"
        cmb4.AddItem "748"
        cmb4.AddItem "749"
    
    ElseIf cmb3 = "750 - 759" Then
    
        cmb4.Clear
        cmb4.AddItem "750"
        cmb4.AddItem "751"
        cmb4.AddItem "752"
        cmb4.AddItem "753"
        cmb4.AddItem "754"
        cmb4.AddItem "755"
        cmb4.AddItem "756"
        cmb4.AddItem "757"
        cmb4.AddItem "758"
        cmb4.AddItem "759"
    
    ElseIf cmb3 = "760 - 769" Then
    
        cmb4.Clear
        cmb4.AddItem "760"
        cmb4.AddItem "761"
        cmb4.AddItem "762"
        cmb4.AddItem "763"
        cmb4.AddItem "764"
        cmb4.AddItem "765"
        cmb4.AddItem "766"
        cmb4.AddItem "767"
        cmb4.AddItem "768"
        cmb4.AddItem "769"
    
    ElseIf cmb3 = "770 - 779" Then
    
        cmb4.Clear
        cmb4.AddItem "770"
        cmb4.AddItem "771"
        cmb4.AddItem "772"
        cmb4.AddItem "773"
        cmb4.AddItem "774"
        cmb4.AddItem "775"
        cmb4.AddItem "776"
        cmb4.AddItem "777"
        cmb4.AddItem "778"
        cmb4.AddItem "779"
    
     ElseIf cmb3 = "780 - 789" Then
     
        cmb4.Clear
        cmb4.AddItem "780"
        cmb4.AddItem "781"
        cmb4.AddItem "782"
        cmb4.AddItem "783"
        cmb4.AddItem "784"
        cmb4.AddItem "785"
        cmb4.AddItem "786"
        cmb4.AddItem "787"
        cmb4.AddItem "788"
        cmb4.AddItem "789"
     
    ElseIf cmb3 = "790 - 799" Then
    
        cmb4.Clear
        cmb4.AddItem "790"
        cmb4.AddItem "791"
        cmb4.AddItem "792"
        cmb4.AddItem "793"
        cmb4.AddItem "794"
        cmb4.AddItem "795"
        cmb4.AddItem "796"
        cmb4.AddItem "797"
        cmb4.AddItem "798"
        cmb4.AddItem "799"
    
    ElseIf cmb3 = "800 - 809" Then
    
        cmb4.Clear
        cmb4.AddItem "800"
        cmb4.AddItem "801"
        cmb4.AddItem "802"
        cmb4.AddItem "803"
        cmb4.AddItem "804"
        cmb4.AddItem "805"
        cmb4.AddItem "806"
        cmb4.AddItem "807"
        cmb4.AddItem "808"
        cmb4.AddItem "809"
    
    ElseIf cmb3 = "810 - 819" Then
    
        cmb4.Clear
        cmb4.AddItem "810"
        cmb4.AddItem "811"
        cmb4.AddItem "812"
        cmb4.AddItem "813"
        cmb4.AddItem "814"
        cmb4.AddItem "815"
        cmb4.AddItem "816"
        cmb4.AddItem "817"
        cmb4.AddItem "818"
        cmb4.AddItem "819"
    
    ElseIf cmb3 = "820 - 829" Then
    
        cmb4.Clear
        cmb4.AddItem "820"
        cmb4.AddItem "821"
        cmb4.AddItem "822"
        cmb4.AddItem "823"
        cmb4.AddItem "824"
        cmb4.AddItem "825"
        cmb4.AddItem "826"
        cmb4.AddItem "827"
        cmb4.AddItem "828"
        cmb4.AddItem "829"
    
    ElseIf cmb3 = "830 - 839" Then
    
        cmb4.Clear
        cmb4.AddItem "830"
        cmb4.AddItem "831"
        cmb4.AddItem "832"
        cmb4.AddItem "833"
        cmb4.AddItem "834"
        cmb4.AddItem "835"
        cmb4.AddItem "836"
        cmb4.AddItem "837"
        cmb4.AddItem "838"
        cmb4.AddItem "839"
    
    ElseIf cmb3 = "840 - 849" Then
    
        cmb4.Clear
        cmb4.AddItem "840"
        cmb4.AddItem "841"
        cmb4.AddItem "842"
        cmb4.AddItem "843"
        cmb4.AddItem "844"
        cmb4.AddItem "845"
        cmb4.AddItem "846"
        cmb4.AddItem "847"
        cmb4.AddItem "848"
        cmb4.AddItem "849"
    
    ElseIf cmb3 = "850 - 859" Then
    
        cmb4.Clear
        cmb4.AddItem "850"
        cmb4.AddItem "851"
        cmb4.AddItem "852"
        cmb4.AddItem "853"
        cmb4.AddItem "854"
        cmb4.AddItem "855"
        cmb4.AddItem "856"
        cmb4.AddItem "857"
        cmb4.AddItem "858"
        cmb4.AddItem "859"
        
    ElseIf cmb3 = "860 - 869" Then
    
        cmb4.Clear
        cmb4.AddItem "860"
        cmb4.AddItem "861"
        cmb4.AddItem "862"
        cmb4.AddItem "863"
        cmb4.AddItem "864"
        cmb4.AddItem "865"
        cmb4.AddItem "866"
        cmb4.AddItem "867"
        cmb4.AddItem "868"
        cmb4.AddItem "869"
    
    ElseIf cmb3 = "870 - 879" Then
    
        cmb4.Clear
        cmb4.AddItem "870"
        cmb4.AddItem "871"
        cmb4.AddItem "872"
        cmb4.AddItem "873"
        cmb4.AddItem "874"
        cmb4.AddItem "875"
        cmb4.AddItem "876"
        cmb4.AddItem "877"
        cmb4.AddItem "878"
        cmb4.AddItem "879"
    
    ElseIf cmb3 = "880 - 889" Then
    
        cmb4.Clear
        cmb4.AddItem "880"
        cmb4.AddItem "881"
        cmb4.AddItem "882"
        cmb4.AddItem "883"
        cmb4.AddItem "884"
        cmb4.AddItem "885"
        cmb4.AddItem "886"
        cmb4.AddItem "887"
        cmb4.AddItem "888"
        cmb4.AddItem "889"
    
    ElseIf cmb3 = "890 - 899" Then
    
        cmb4.Clear
        cmb4.AddItem "890"
        cmb4.AddItem "891"
        cmb4.AddItem "892"
        cmb4.AddItem "893"
        cmb4.AddItem "894"
        cmb4.AddItem "895"
        cmb4.AddItem "896"
        cmb4.AddItem "897"
        cmb4.AddItem "898"
        cmb4.AddItem "899"
    
    ElseIf cmb3 = "900 - 909" Then
    
        cmb4.Clear
        cmb4.AddItem "900"
        cmb4.AddItem "901"
        cmb4.AddItem "902"
        cmb4.AddItem "903"
        cmb4.AddItem "904"
        cmb4.AddItem "905"
        cmb4.AddItem "906"
        cmb4.AddItem "907"
        cmb4.AddItem "908"
        cmb4.AddItem "909"
    
    ElseIf cmb3 = "910 - 919" Then
    
        cmb4.Clear
        cmb4.AddItem "910"
        cmb4.AddItem "911"
        cmb4.AddItem "912"
        cmb4.AddItem "913"
        cmb4.AddItem "914"
        cmb4.AddItem "915"
        cmb4.AddItem "916"
        cmb4.AddItem "917"
        cmb4.AddItem "918"
        cmb4.AddItem "919"
    
    ElseIf cmb3 = "920 - 929" Then
    
        cmb4.Clear
        cmb4.AddItem "920"
        cmb4.AddItem "921"
        cmb4.AddItem "922"
        cmb4.AddItem "923"
        cmb4.AddItem "924"
        cmb4.AddItem "925"
        cmb4.AddItem "926"
        cmb4.AddItem "927"
        cmb4.AddItem "928"
        cmb4.AddItem "929"
    
    ElseIf cmb3 = "930 - 939" Then
    
        cmb4.Clear
        cmb4.AddItem "930"
        cmb4.AddItem "931"
        cmb4.AddItem "932"
        cmb4.AddItem "933"
        cmb4.AddItem "934"
        cmb4.AddItem "935"
        cmb4.AddItem "936"
        cmb4.AddItem "937"
        cmb4.AddItem "938"
        cmb4.AddItem "939"
    
    ElseIf cmb3 = "940 - 949" Then
    
        cmb4.Clear
        cmb4.AddItem "940"
        cmb4.AddItem "941"
        cmb4.AddItem "942"
        cmb4.AddItem "943"
        cmb4.AddItem "944"
        cmb4.AddItem "945"
        cmb4.AddItem "946"
        cmb4.AddItem "947"
        cmb4.AddItem "948"
        cmb4.AddItem "949"
    
    ElseIf cmb3 = "950 - 959" Then
    
        cmb4.Clear
        cmb4.AddItem "950"
        cmb4.AddItem "951"
        cmb4.AddItem "952"
        cmb4.AddItem "953"
        cmb4.AddItem "954"
        cmb4.AddItem "955"
        cmb4.AddItem "956"
        cmb4.AddItem "957"
        cmb4.AddItem "958"
        cmb4.AddItem "959"
    
    ElseIf cmb3 = "960 - 969" Then
    
        cmb4.Clear
        cmb4.AddItem "960"
        cmb4.AddItem "961"
        cmb4.AddItem "962"
        cmb4.AddItem "963"
        cmb4.AddItem "964"
        cmb4.AddItem "965"
        cmb4.AddItem "966"
        cmb4.AddItem "967"
        cmb4.AddItem "968"
        cmb4.AddItem "969"
    
    ElseIf cmb3 = "970 - 979" Then
    
        cmb4.Clear
        cmb4.AddItem "970"
        cmb4.AddItem "971"
        cmb4.AddItem "972"
        cmb4.AddItem "973"
        cmb4.AddItem "974"
        cmb4.AddItem "975"
        cmb4.AddItem "976"
        cmb4.AddItem "977"
        cmb4.AddItem "978"
        cmb4.AddItem "979"
    
    ElseIf cmb3 = "980 - 989" Then
    
        cmb4.Clear
        cmb4.AddItem "980"
        cmb4.AddItem "981"
        cmb4.AddItem "982"
        cmb4.AddItem "983"
        cmb4.AddItem "984"
        cmb4.AddItem "985"
        cmb4.AddItem "986"
        cmb4.AddItem "987"
        cmb4.AddItem "988"
        cmb4.AddItem "989"
    
    ElseIf cmb3 = "990 - 999" Then
    
        cmb4.Clear
        cmb4.AddItem "990"
        cmb4.AddItem "991"
        cmb4.AddItem "992"
        cmb4.AddItem "993"
        cmb4.AddItem "994"
        cmb4.AddItem "995"
        cmb4.AddItem "996"
        cmb4.AddItem "997"
        cmb4.AddItem "998"
        cmb4.AddItem "999"
    
    ElseIf cmb3 = "1000 - 1009" Then
    
        cmb4.Clear
        cmb4.AddItem "1000"
        cmb4.AddItem "1001"
        cmb4.AddItem "1002"
        cmb4.AddItem "1003"
        cmb4.AddItem "1004"
        cmb4.AddItem "1005"
        cmb4.AddItem "1006"
        cmb4.AddItem "1007"
        cmb4.AddItem "1008"
        cmb4.AddItem "1009"
    
        cmb3 = "1010 - 1019"
        cmb3 = "1020 - 1029"
        cmb3 = "1030 - 1039"
        cmb3 = "1040 - 1049"
        cmb3 = "1050 - 1059"
        cmb3 = "1060 - 1069"
        cmb3 = "1070 - 1079"
        cmb3 = "1080 - 1089"
        cmb3 = "1090 - 1099"
        cmb3 = "1100 - 1109"
        cmb3 = "1110 - 1119"
        cmb3 = "1120 - 1129"
        cmb3 = "1130 - 1139"
        cmb3 = "1140 - 1149"
        cmb3 = "1150 - 1159"
        cmb3 = "1160 - 1169"
        cmb3 = "1170 - 1179"
        cmb3 = "1180 - 1189"
        cmb3 = "1190 - 1199"
        cmb3 = "1200 - 1209"
        cmb3 = "1210 - 1219"
        cmb3 = "1220 - 1229"
        cmb3 = "1230 - 1239"
        cmb3 = "1240 - 1249"
        cmb3 = "1250 - 1259"
        cmb3 = "1260 - 1269"
        cmb3 = "1270 - 1279"
        cmb3 = "1280 - 1289"
        cmb3 = "1290 - 1299"
        cmb3 = "1300 - 1309"
        cmb3 = "1310 - 1319"
        cmb3 = "1320 - 1329"
        cmb3 = "1330 - 1339"
        cmb3 = "1340 - 1349"
        cmb3 = "1350 - 1359"
        cmb3 = "1360 - 1369"
        cmb3 = "1370 - 1379"
        cmb3 = "1380 - 1389"
        cmb3 = "1390 - 1399"
        cmb3 = "1400 - 1409"
        cmb3 = "1410 - 1419"
        cmb3 = "1420 - 1429"
        cmb3 = "1430 - 1439"
        cmb3 = "1440 - 1449"
        cmb3 = "1450 - 1459"
        cmb3 = "1460 - 1469"
        cmb3 = "1470 - 1479"
        cmb3 = "1480 - 1489"
        cmb3 = "1490 - 1499"
        cmb3 = "1500 - 1509"
        cmb3 = "1510 - 1519"
        cmb3 = "1520 - 1529"
        cmb3 = "1530 - 1539"
        cmb3 = "1540 - 1549"
        cmb3 = "1550 - 1559"
        cmb3 = "1560 - 1569"
        cmb3 = "1570 - 1579"
        cmb3 = "1580 - 1589"
        cmb3 = "1590 - 1599"
        cmb3 = "1600 - 1609"
        cmb3 = "1610 - 1619"
        cmb3 = "1620 - 1629"
        cmb3 = "1630 - 1639"
        cmb3 = "1640 - 1649"
        cmb3 = "1650 - 1659"
        cmb3 = "1660 - 1669"
        cmb3 = "1670 - 1679"
        cmb3 = "1680 - 1689"
        cmb3 = "1690 - 1699"
        cmb3 = "1700 - 1709"
        cmb3 = "1710 - 1719"
        cmb3 = "1720 - 1729"
        cmb3 = "1730 - 1739"
        cmb3 = "1740 - 1749"
        cmb3 = "1750 - 1759"
        cmb3 = "1760 - 1769"
        cmb3 = "1770 - 1779"
        cmb3 = "1780 - 1789"
        cmb3 = "1790 - 1799"
        cmb3 = "1800 - 1809"
        cmb3 = "1810 - 1819"
        cmb3 = "1820 - 1829"
        cmb3 = "1830 - 1839"
        cmb3 = "1840 - 1849"
        cmb3 = "1850 - 1859"
        cmb3 = "1860 - 1869"
        cmb3 = "1870 - 1879"
        cmb3 = "1880 - 1889"
        cmb3 = "1890 - 1899"
        cmb3 = "1900 - 1909"
        cmb3 = "1910 - 1919"
        cmb3 = "1920 - 1929"
        cmb3 = "1930 - 1939"
        cmb3 = "1940 - 1949"
        cmb3 = "1950 - 1959"
        cmb3 = "1960 - 1969"
        cmb3 = "1970 - 1979"
        cmb3 = "1980 - 1989"
        cmb3 = "1990 - 1999"
        cmb3 = "2000 - 2009"
        cmb3 = "2010 - 2019"
        cmb3 = "2020 - 2029"
        cmb3 = "2030 - 2039"
        cmb3 = "2040 - 2049"
        cmb3 = "2050 - 2059"
        cmb3 = "2060 - 2069"
        cmb3 = "2070 - 2079"
        cmb3 = "2080 - 2089"
        cmb3 = "2090 - 2099"
        cmb3 = "2100 - 2109"
        cmb3 = "2110 - 2119"
        cmb3 = "2120 - 2129"
        cmb3 = "2130 - 2139"
        cmb3 = "2140 - 2149"
        cmb3 = "2150 - 2159"
        cmb3 = "2160 - 2169"
        cmb3 = "2170 - 2179"
        cmb3 = "2180 - 2189"
        cmb3 = "2190 - 2199"
        cmb3 = "2200 - 2209"
        cmb3 = "2210 - 2219"
        cmb3 = "2220 - 2224"

    End If
End Sub

Private Sub cmb4_Click()

    frmMain.txtFOB = cmb4.Text
    frmMain.Show
    frmFuel.Hide

End Sub

Private Sub Form_Activate()

    cmb1 = ""
    cmb2.Clear
    cmb3.Clear
    cmb4.Clear
    cmb1.SetFocus

End Sub

Private Sub Form_Load()
    
    cmb1.AddItem "35 - 999"
    cmb1.AddItem "1000 - 1499"
    cmb1.AddItem "1500 - 2224"
    
    cmb2.Clear
    cmb3.Clear
    cmb4.Clear
    
    cmb4.AddItem "35"
    cmb4.AddItem "36"

End Sub
