VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Multiline Tooltip Gaya Balon"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Source Code dimulai dari sini

Option Explicit
'Pemanggilan fungsi API diperlukan untuk membuat dan 'menghancurkan tooltip di Sistem Operasi Windows.

Private Declare Function CreateWindowEx Lib "user32" _
Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
ByVal lpClassName As String, ByVal lpWindowName As _
String, ByVal dwStyle As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight _
As Long, ByVal hWndParent As Long, ByVal hMenu _
As Long, ByVal hInstance As Long, lpParam As Any) _
As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Function GetClientRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function DestroyWindow Lib "user32" _
(ByVal hwnd As Long) As Long

'UDT (User Defined Type) RECT.
'Digunakan untuk pengaturan batas dari jendela tooltip.
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'UDT TOOLINFO.
'Digunakan untuk menentukan semua tanda yang diperlukan
'untuk membuat jendela tooltip.
Private Type TOOLINFO
  cbSize As Long
  uFlags As Long
  hwnd As Long
  uid As Long
  RECT As RECT
  hinst As Long
  lpszText As String
  lParam As Long
End Type

'Sebuah konstanta yang digunakan untuk menghubungkan
'ke fungsi API yang bernama: CreateWindowEx.
'Hal ini untuk menandakan nilai default yang digunakan.
Private Const CW_USEDEFAULT = &H80000000

'Konstanta untuk fungsi API bernama: SetWindowPosition.
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1

'Konstanta untuk menentukan gaya dari jendela tooltip.
Private Const WS_POPUP = &H80000000
Private Const WS_EX_TOPMOST = &H8&

'Konstanta yang digunakan dengan fungsi API SendMessage
'untuk mendefinisikan pesan private.
Private Const WM_USER = &H400

'Messages yang digunakan untuk menentukan durasi waktu 'dari tooltips. Tidak digunakan di sini.
Private Const TTDT_AUTOMATIC = 0
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTDT_RESHOW = 1

'Semua "penanda" untuk jendela tooltip.
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_CENTERTIP = &H2
Private Const TTF_DI_SETITEM = &H8000
Private Const TTF_IDISHWND = &H1
Private Const TTF_RTLREADING = &H4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_TRANSPARENT = &H100

'Semua pesan yang tersedia untuk tooltip Windows.
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ADDTOOLW = (WM_USER + 50)
Private Const TTM_ADJUSTRECT = (WM_USER + 31)
Private Const TTM_DELTOOLA = (WM_USER + 5)
Private Const TTM_DELTOOLW = (WM_USER + 51)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETBUBBLESIZE = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
Private Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
Private Const TTM_GETDELAYTIME = (WM_USER + 21)
Private Const TTM_GETMARGIN = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_GETTEXTW = (WM_USER + 56)
Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_GETTOOLINFOA = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW = (WM_USER + 53)
Private Const TTM_HITTESTA = (WM_USER + 10)
Private Const TTM_HITTESTW = (WM_USER + 55)
Private Const TTM_NEWTOOLRECTA = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW = (WM_USER + 52)
Private Const TTM_POP = (WM_USER + 28)
Private Const TTM_RELAYEVENT = (WM_USER + 7)
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_SETMARGIN = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLEA = (WM_USER + 32)
Private Const TTM_SETTITLEW = (WM_USER + 33)
Private Const TTM_SETTOOLINFOA = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
Private Const TTM_TRACKACTIVATE = (WM_USER + 17)
Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_UPDATE = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16)

'Konstanta untuk menentukan gaya dari jendela tooltip.
'Selalu tip, walalupun jika jendela utama tidak aktif.
Private Const TTS_ALWAYSTIP = &H1
'Menggunakan gaya balon tooltip.
Private Const TTS_BALLOON = &H40
'Win98 and up - jangan gunakan sliding tooltips.
Private Const TTS_NOANIMATE = &H10
'Win2K and up - jangan hilangkan tooltips.
Private Const TTS_NOFADE = &H20
'Mencegah Windows dari penghapusan karakter ampersand 'apapun di dalam string tooltip. Tanpa penanda ini, 'Windows otomatis akan menghapus karakter ampersand 'dari string tersebut. Hal ini dilakukan untuk 'mengizinkan string yang sama dapat digunakan
'sebagai teks dari tooltip, dan sebagai tulisan dari 'sebuah control.
Private Const TTS_NOPREFIX = &H2

'Class untuk dua tooltip yang berbeda.
Private Const TOOLTIPS_CLASS = "tooltips_class"
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'Sebuah variabel bertipe Long untuk menyimpan hwnd '(window handle) dari jendela tooltip yang dibuat di 'contoh ini.Hal ini akan menjadi sebuah array bertipe 'Long jika kita membuat tooltip Windows untuk banyak 'control atau banyak jendela.
Dim hwndTT As Long
 
'Event Code. Untuk mencoba coding ini, yakinkan sekali 'lagi bahwa di form Anda sudah ada 2 tombol bernama "Command1" dan
'"Command2"

Private Sub Form_Load()
  'Deklarasikan sebuah variabel bertipe UDT TOOLINFO.
  Dim ti As TOOLINFO

  'Variabel ini digunakan untuk menandakan batas dari
  'jendela tooltip
  Dim RECT As RECT

  'Untuk melewatkan toolinfo UDT sebagai sebuah ID
  'untuk jendela tooltip. Tidak melakukan apapun di
  'contoh ini, untuk menjelaskan saja.
  Dim uid As Long
  uid = 0

  'Sebuah string yang akan ditampilkan di dalam
  'tooltip.
  Dim strPntr As String
  strPntr = "Inilah tooltip yang dibuat dengan menggunakan fungsi API. " & vbCrLf & "Seperti yang dapat Anda lihat, dia kini mendukung banyak baris, " & vbCrLf & _
  "pindah baris, menampilkan batas atau jendela tooltip bergaya balon, " & vbCrLf & _
  "serta dapat menampilkan warna latar dan huruf sesuai keinginan."

  'Nilai yang dikembalikan saat pemanggilan fungsi API.
  Dim RetVal As Long

  'Buat sebuah jendela tooltip, dan tangani hwnd-nya di
  'dalam lebar form hwndTT yang bertipe Long.
  hwndTT = CreateWindowEx(WS_EX_TOPMOST, _
           TOOLTIPS_CLASSA, vbNullString, _
           WS_POPUP Or TTS_NOPREFIX Or TTS_BALLOON, _
           CW_USEDEFAULT, CW_USEDEFAULT, _
           CW_USEDEFAULT, CW_USEDEFAULT, _
           Me.hwnd, 0, App.hInstance, 0)
  'Gunakan fungsi API setwindowpos untuk menentukan
  'posisi jendela dari tooltip.
  SetWindowPos hwndTT, HWND_TOPMOST, 0, 0, 0, 0, _
          SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

  'Mendeteksi batas control yang tooltipnya sedang
  'ditambahkan. Ini akan menjadi batas untuk
  'mengaktifkan jendela tooltip.
  GetClientRect Command1.hwnd, RECT

  'Tentukan semua informasi yang diperlukan untuk
  'toolinfo UDT.
  
  'Ukuran UDT toolinfo dalam bytes. Harus di-set!
  ti.cbSize = Len(ti)
  
  'Penanda yang akan kita lewatkan untuk tooltip.
  'TTF_CENTERTIP tidak perlu, tapi tengahkan tooltip ke
  'jendela di mana tooltip sedang diaplikasikan (jika
  'memungkinkan). TTF_SUBCLASS memberitahukan ke
  'jendela tooltip window untuk meng-sub-class jendela
  'yang sedang diaplikasikan. Ini cara terbaik di VB,
  'jadi subclassing oleh pengembang tidak diperlukan.
  ti.uFlags = TTF_CENTERTIP Or TTF_SUBCLASS
  
  'hwnd dari control yang tooltipnya sedang
  'diaplikasikan.
  ti.hwnd = Command1.hwnd
  
  'Instansiasi dari aplikasi yang tooltip-nya sedang
  'diaplikasikan.
  ti.hinst = App.hInstance
  'ID (hwnd) dari jendela tooltip. Tidak diperlukan
  'kecuali jendela dibuat dengan menggunakan penanda TTF_IDISHWND.
  ti.uid = uid
  
  'Sebuah pointer ke tooltip.
  ti.lpszText = strPntr
  
  'Koordinat yang menentukan batas jendela tooltip
  'ketika aktif.
  ti.RECT.Left = RECT.Left
  ti.RECT.Top = RECT.Top
  ti.RECT.Right = RECT.Right
  ti.RECT.Bottom = RECT.Bottom

  'Kirim sebuah pesan ke jendela tooltip untuk
  'menampilkan tooltip pada control yang sedang
  'diaplikasikan.
  RetVal = SendMessage(hwndTT, TTM_ADDTOOLA, 0, ti)

  'Kirim sebuah pesan ke jendela tooltip untuk
  'menentukan lebar maksimum agar dapat mendukung
  'pindah baris (line-breaking).
  RetVal = SendMessage(hwndTT, TTM_SETMAXTIPWIDTH, _
           0, 80)

  'Kirim pesan ke jendela tooltip untuk menentukan
  'warna latar balon dan warna huruf. Dalam hal ini,
  'kita menggunakan fungsi warna RGB
  'RetVal = SendMessage(hwndTT, TTM_SETTIPBKCOLOR, _
  '          RGB(255, 255, 255), 0)
  'Coba ganti warna latar dengan hijau muda...
  RetVal = SendMessage(hwndTT, TTM_SETTIPBKCOLOR, _
           &HC0FFC0, 0)
  'RetVal = SendMessage(hwndTT, TTM_SETTIPTEXTCOLOR,
            'RGB(0, 0, 150), 0)
  'Coba ganti warna huruf tooltip dengan warna biru
  RetVal = SendMessage(hwndTT, TTM_SETTIPTEXTCOLOR, _
           vbBlue, 0)

  'Kirim sebuah pesan ke jendela tooltip untuk
  'mengupdate dirinya.
  '(Jika ada warna latar dan huruf yang baru).
  RetVal = SendMessage(hwndTT, TTM_UPDATETIPTEXTA, 0, ti)

  'Tentukan teks dari tombol kedua untuk menampilkan
  'tooltip standar milik Visual Basic (tidak mendukung
  'multi-line).
  Command2.ToolTipText = "Inilah tooltip standar VB." & vbCrLf & _
  "Seperti yang Anda lihat, karakter CrLf di sebelah kiri " & "baris ini tidak berfungsi di sini. " & _
  vbCrLf & "Karakter VbCrLf ditandai dengan garis dua tebal vertikal"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Ketika form unload, yakinkan bahwa tooltip yang
  'dibuat dihancurkan (dibebaskan dari memory)
  DestroyWindow hwndTT
End Sub


