VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_Zeit 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Zeit abgleichen"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   1815
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Timer Timer_ende 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   120
   End
   Begin InetCtlsObjects.Inet Inet_Zeit 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_Zeit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Ausführen beim starten des Programms
Call zeitabgleich
End Sub

Private Sub zeitabgleich()
'Variablen Deklaration
Dim Datum_Uhrzeit, Quelle As String
Dim Datum, Uhrzeit As String
Dim Kanal, i As Integer

'Bei Fehler nächste Zeile
On Error Resume Next

For i = 0 To 3
    Select Case i
    Case 0
    'Ersatz Quelle
    Quelle = "http://erassoft.er.funpic.de/zeit.php"
    Case 1
    'Standardisierte Quelle
    Quelle = "http://erassoft.de/zeit.php"
    Case 2
    'Sub Quelle
    Quelle = "http://time.erassoft.de"
    Case 3
    'andere Quelle der PHP-Uhrzeit nutzen?
    If Dir("zeitabgleich.txt") <> "" Then
    'Textdatei mit Namen "zeitabgleich.txt" mit der URL lesen
    Kanal = FreeFile
    Open ("zeitabgleich.txt") For Input As #Kanal  'hier wird die Datei geöffnet
    Input #1, Quelle
    Close #Kanal
    End If
    End Select
    
'Online Datum/Uhrzeit auslesen
Datum_Uhrzeit = Inet_Zeit.OpenURL(Quelle)
'Online Datum/Uhrzeit in Variablen speichern
Datum = Mid(Datum_Uhrzeit, 1, 10)
Uhrzeit = Mid(Datum_Uhrzeit, 12, 8)
'Online Datum/Uhrzeit in Computeruhr übertragen
Date = (Datum)                'zB. "24.08.1990"
Time = (Uhrzeit)              'zB. "12:00:00"
Next i

'Timer zum Programm beenden starten
Timer_ende.Interval = 10
Timer_ende.Enabled = True
End Sub

Private Sub Timer_ende_Timer()
'Programm beenden
End
End Sub
