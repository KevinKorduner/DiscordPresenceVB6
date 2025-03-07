VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Discord"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "Creado por Kevin Korduner"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Github: https://github.com/KevinKorduner/
' Creado y liberado por Kevin Alexis Korduner 07/03/2025
' Este ejecutable hace el llamado a la DLL, deben configurar su app ID de discord.

' ==============================================================
' Declaraciones globales y de tipos
' ==============================================================

' Variables globales para almacenar los punteros a las cadenas ANSI.
Private g_statePtr As Long
Private g_detailsPtr As Long

' --------------------------------------------------------------
' Estructura para los callbacks (no se usan en este ejemplo)
' --------------------------------------------------------------
Private Type DiscordEventHandlers
    ready As Long           ' puntero a función (no se usa)
    disconnected As Long
    errored As Long
    joinGame As Long
    spectateGame As Long
    joinRequest As Long
End Type

' --------------------------------------------------------------
' Estructura adaptada de Discord Rich Presence
' (Los timestamps se dividen en dos Long: parte baja y parte alta)
' --------------------------------------------------------------
Private Type DiscordRichPresence
    state As Long             ' puntero (char*) a cadena ANSI
    details As Long           ' puntero (char*) a cadena ANSI
    startTimestampLow As Long ' parte baja del timestamp (int64)
    startTimestampHigh As Long ' parte alta del timestamp
    endTimestampLow As Long
    endTimestampHigh As Long
    largeImageKey As Long     ' puntero (char*) a cadena ANSI
    largeImageText As Long    ' puntero (char*) a cadena ANSI
    smallImageKey As Long     ' puntero (char*) a cadena ANSI
    smallImageText As Long    ' puntero (char*) a cadena ANSI
    partyId As Long           ' puntero (char*) a cadena ANSI
    partySize As Long         ' entero
    partyMax As Long          ' entero
    matchSecret As Long       ' puntero (char*) a cadena ANSI
    joinSecret As Long        ' puntero (char*) a cadena ANSI
    spectateSecret As Long    ' puntero (char*) a cadena ANSI
    instance As Long          ' entero
End Type

' --------------------------------------------------------------
' Declaración de funciones exportadas por discord-rpc.dll
' --------------------------------------------------------------
Private Declare Sub Discord_Initialize Lib "discord-rpc.dll" ( _
    ByVal applicationId As String, _
    ByRef handlers As DiscordEventHandlers, _
    ByVal autoRegister As Long, _
    ByVal optionalSteamId As String)

Private Declare Sub Discord_UpdatePresence Lib "discord-rpc.dll" ( _
    ByRef presence As DiscordRichPresence)

Private Declare Sub Discord_RunCallbacks Lib "discord-rpc.dll" ()
Private Declare Sub Discord_Shutdown Lib "discord-rpc.dll" ()

' --------------------------------------------------------------
' Declaraciones de funciones API para manejo de memoria
' --------------------------------------------------------------
Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Long, Source As Long, ByVal Length As Long)

Private Const GMEM_FIXED As Long = &H0

' --------------------------------------------------------------
' Función para asignar memoria y crear una cadena ANSI
' --------------------------------------------------------------
Public Function AllocAnsiString(s As String) As Long
    Dim byteArray() As Byte
    Dim lenBytes As Long
    Dim p As Long

    ' Convierte la cadena de Unicode a ANSI y añade el terminador nulo.
    byteArray = StrConv(s & vbNullChar, vbFromUnicode)
    lenBytes = UBound(byteArray) - LBound(byteArray) + 1

    p = GlobalAlloc(GMEM_FIXED, lenBytes)
    If p <> 0 Then
        CopyMemory p, VarPtr(byteArray(0)), lenBytes
    Else
        MsgBox "Error al asignar memoria para la cadena: " & s, vbCritical, "Error"
    End If
    AllocAnsiString = p
End Function

' --------------------------------------------------------------
' Procedimiento para liberar la memoria asignada
' --------------------------------------------------------------
Public Sub FreeAnsiString(p As Long)
    If p <> 0 Then GlobalFree p
End Sub

' ==============================================================
' Eventos del formulario
' ==============================================================

Private Sub Form_Load()
    ' Llamada a la subrutina de inicialización al cargar el formulario.
    Ta
End Sub

Private Sub Ta()
    On Error GoTo ErrHandler
    
    Dim sDllPath As String
    sDllPath = App.Path
    ' Asegurarse de que la ruta termine en "\"
    If Right$(sDllPath, 1) <> "\" Then
        sDllPath = sDllPath & "\"
    End If
    
    ' Comprueba que la DLL exista en el directorio de la aplicación.
    If Dir$(sDllPath & "discord-rpc.dll") = "" Then
        MsgBox "No se encontró discord-rpc.dll en " & sDllPath, vbCritical, "Error"
        Exit Sub
    End If
    
    Dim handlers As DiscordEventHandlers
    Dim presence As DiscordRichPresence
    Dim unixTime As Long
    
    Debug.Print "Inicializando Discord RPC..."
    
    ' Inicializa la estructura de callbacks a cero (no se usan en este ejemplo)
    handlers.ready = 0
    handlers.disconnected = 0
    handlers.errored = 0
    handlers.joinGame = 0
    handlers.spectateGame = 0
    handlers.joinRequest = 0

    ' Llama a Discord_Initialize con tu App ID.
    ' (La DLL debe gestionar el terminador nulo de la cadena)
    Discord_Initialize "5234441513530066311", handlers, 1, ""
    Debug.Print "Discord_Initialize ejecutado."
    
    ' Asigna memoria y prepara las cadenas ANSI para "state" y "details"
    g_statePtr = AllocAnsiString("Jugando AOForever")
    g_detailsPtr = AllocAnsiString("Jugando AOForever")
    If g_statePtr = 0 Or g_detailsPtr = 0 Then
        MsgBox "Error al asignar memoria para las cadenas ANSI.", vbCritical, "Error"
        Exit Sub
    End If
    
    presence.state = g_statePtr
    presence.details = g_detailsPtr
    
    ' Calcula el timestamp UNIX (segundos desde el 1/1/1970)
    unixTime = DateDiff("s", "1/1/1970 00:00:00", Now)
    presence.startTimestampLow = unixTime
    presence.startTimestampHigh = 0
    presence.endTimestampLow = 0
    presence.endTimestampHigh = 0
    
    ' Para este ejemplo no se usan imágenes, party, secretos ni instancia.
    presence.largeImageKey = 0
    presence.largeImageText = 0
    presence.smallImageKey = 0
    presence.smallImageText = 0
    presence.partyId = 0
    presence.partySize = 0
    presence.partyMax = 0
    presence.matchSecret = 0
    presence.joinSecret = 0
    presence.spectateSecret = 0
    presence.instance = 0
    
    ' Actualiza la presencia en Discord
    Discord_UpdatePresence presence
    Debug.Print "Discord_UpdatePresence ejecutado."
    
    ' Configura y activa el Timer para llamar a Discord_RunCallbacks cada segundo
    Timer1.Interval = 1000   ' 1000 ms = 1 segundo
    Timer1.Enabled = True
    
    ' Actualiza un Label para indicar que la aplicación está en ejecución.
    Label1.Caption = "Inicialización completa. La aplicación se encuentra en ejecución." & vbCrLf & _
                     "Cierre el formulario para salir y retirar el tag de Discord."
    Debug.Print "Inicialización completa."
    
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    ' Llama a los callbacks de Discord para mantener la conexión.
    Discord_RunCallbacks
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Al cerrar el formulario: detiene el Timer, cierra la conexión y libera la memoria.
    Timer1.Enabled = False
    Discord_Shutdown
    FreeAnsiString g_statePtr
    FreeAnsiString g_detailsPtr
End Sub


