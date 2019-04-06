VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de censura"
   ClientHeight    =   6825
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5925
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboDevices 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Início e fim Agendados"
      Height          =   1095
      Left            =   0
      TabIndex        =   25
      Top             =   4920
      Width           =   5895
      Begin VB.CheckBox CheckProgramado 
         Caption         =   "Habilitar"
         Height          =   285
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TextFim 
         Height          =   285
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   27
         Text            =   "00:00:00"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TextIni 
         Height          =   285
         Left            =   3120
         MaxLength       =   8
         TabIndex        =   26
         Text            =   "00:00:00"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label LabelRelogio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label LabelFim 
         Caption         =   "Fim"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LabelIni 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Gravação"
      TabPicture(0)   =   "Main.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelNivel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblInputType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelInfos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Labelhora"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LabelClipL"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LabelClipR"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TimerGravacao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ProgressBarLeft"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ProgressBarRight"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ComboQualidade"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dialogo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TimerSegundos"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TimerEventos"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TimerRelogio"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Biblioteca"
      TabPicture(1)   =   "Main.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LabelDiretorio"
      Tab(1).Control(1)=   "TextDiretorio"
      Tab(1).Control(2)=   "CommandDiretorio"
      Tab(1).Control(3)=   "ListViewAudios"
      Tab(1).Control(4)=   "CommandPlay"
      Tab(1).Control(5)=   "CommandStop"
      Tab(1).ControlCount=   6
      Begin VB.Timer TimerRelogio 
         Interval        =   1000
         Left            =   3840
         Top             =   4200
      End
      Begin VB.Timer TimerEventos 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4320
         Top             =   4200
      End
      Begin VB.CommandButton CommandStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   -70200
         TabIndex        =   21
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton CommandPlay 
         Caption         =   "Play"
         Height          =   375
         Left            =   -71280
         TabIndex        =   20
         Top             =   4200
         Width           =   975
      End
      Begin VB.Timer TimerSegundos 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4800
         Top             =   4200
      End
      Begin MSComDlg.CommonDialog dialogo 
         Left            =   1920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListViewAudios 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5318
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nome"
            Object.Width           =   9438
         EndProperty
      End
      Begin VB.CommandButton CommandDiretorio 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   -70080
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TextDiretorio 
         Height          =   285
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   4695
      End
      Begin VB.ComboBox ComboQualidade 
         Height          =   315
         ItemData        =   "Main.frx":688A
         Left            =   240
         List            =   "Main.frx":68AB
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5655
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   3240
            Picture         =   "Main.frx":68EA
            ScaleHeight     =   1095
            ScaleWidth      =   2295
            TabIndex        =   18
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox ComboDivisao 
            Height          =   315
            ItemData        =   "Main.frx":C754
            Left            =   120
            List            =   "Main.frx":C782
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1920
            Width           =   2895
         End
         Begin VB.CommandButton btnRecord 
            Height          =   345
            Left            =   3840
            Picture         =   "Main.frx":C7CA
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1920
            Width           =   1095
         End
         Begin VB.ComboBox cmbInput 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   480
            Width           =   2895
         End
         Begin MSComctlLib.Slider sldInputLevel 
            Height          =   255
            Left            =   3240
            TabIndex        =   7
            Top             =   1560
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
            SelectRange     =   -1  'True
            TickFrequency   =   10
         End
         Begin VB.Label LabelDuracao 
            Caption         =   "Dividr audio a cada:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label LabelQuality 
            Caption         =   "Qualidade da Gravação"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label LabelCanal 
            Caption         =   "Canal de gravação"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBarRight 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   3720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Max             =   32767
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBarLeft 
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   3240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Max             =   32767
         Scrolling       =   1
      End
      Begin VB.Timer TimerGravacao 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5280
         Top             =   4200
      End
      Begin VB.Label LabelClipR 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   5520
         TabIndex        =   24
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label LabelClipL 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   5520
         TabIndex        =   23
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Labelhora 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label LabelInfos 
         BackColor       =   &H00C0C0C0&
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label LabelDiretorio 
         Caption         =   "Pasta de trabalho"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblInputType 
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   3000
         Width           =   1680
      End
      Begin VB.Label LabelNivel 
         Caption         =   "Nível de entrada -"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   3000
         Width           =   1335
      End
   End
   Begin VB.Label LabelDevice 
      Caption         =   "Interface de Gravação"
      Height          =   255
      Left            =   2880
      TabIndex        =   33
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnusobre 
      Caption         =   "Sobre"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'niveis de gravação
Dim lowPal As Long
Dim hiPal As Long
'arquivos
Dim myFSO As New Scripting.FileSystemObject
Dim diretorio As Scripting.Folder
Dim disco As Scripting.Drive
Dim Contador As Integer

Private Sub CheckProgramado_Click()
 'verificando se a sintaxe está correta
    If (CheckProgramado.value = 1) Then
        Dim objRegExp As RegExp
        Set objRegExp = New RegExp
    
        'testando a sintaxe da programação
        objRegExp.Pattern = "(([0-2][0-9]){1}:([0-5][0-9]){1}:([0-5][0-9]){1})"
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        
        If (objRegExp.Test(TextIni.Text) = True And objRegExp.Test(TextFim.Text) = True) Then
            If (Not diretorio Is Nothing) Then
                TextIni.Enabled = False
                TextFim.Enabled = False
                TimerEventos.Enabled = True
            Else
                MsgBox "A pasta de trabalho ainda não foi selecionada!"
                CheckProgramado.value = 0
            End If
        Else
            MsgBox ("Os Horários estão com erro de sintaxe")
            CheckProgramado.value = 0
        End If
    Else
        TimerEventos.Enabled = False
        TextIni.Enabled = True
        TextFim.Enabled = True
    End If
End Sub

Private Sub ComboDevices_Change()
    Call reloadDeviceInputs
End Sub

'escolhe diretório de trabalho
Private Sub CommandDiretorio_Click()
    Dim objShell As New Shell32.Shell
    Dim objFolder As Shell32.Folder2
    Set objFolder = objShell.BrowseForFolder(Me.hWnd, "Select a Folder", 0, ssfDRIVES)

    If Not objFolder Is Nothing Then
        TextDiretorio.Text = objFolder.Self.Path
        Call recuperaArquivos(objFolder.Self.Path)
    End If
    Set objFolder = Nothing
End Sub

Private Sub Form_Load()
    'selecionando a pasta do programa para carregar as dlls
    ChDrive App.Path
    ChDir App.Path
    ComboDivisao.ListIndex = 0
    ComboQualidade.ListIndex = 1
    'informado que não tem nenhum evento programado em andamento
    'verificando versão do bass.dll
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("Uma versão incorreta da Bass.dll foi carregada.", vbCritical)
        End
    End If
    '////////////////////////////////////////////////////////////////////////////////////////////
    'CARREGANDO DISPOSITIVOS
    Dim cnt As Integer
    Dim i As BASS_DEVICEINFO
    cnt = 1      ' device 1 = 1st real device
    While BASS_GetDeviceInfo(cnt, i)
        If (i.flags And BASS_DEVICE_ENABLED) Then  ' enabled, so add it...
            ComboDevices.AddItem VBStrFromAnsiPtr(i.name)
            ComboDevices.ItemData(ComboDevices.NewIndex) = cnt    'store device #
        End If
        cnt = cnt + 1
    Wend
    '////////////////////////////////////////////////////////////////////////////////////////////
    If (ComboDevices.ListCount > 0) Then
        ComboDevices.ListIndex = 0
        'selecionando dispositivo de gravação
        If (BASS_RecordInit(ComboDevices.ListIndex) = 0) Or (BASS_Init(ComboDevices.ListIndex, 44100, 0, Me.hWnd, 0) = 0) Then
            Call Error_("O dispositivo não pode ser iniciado")
            End
        Else
            'carregando uma lista de entradas
            Dim c As Integer
            input_ = -1
            While BASS_RecordGetInputName(c)
                cmbInput.AddItem VBStrFromAnsiPtr(BASS_RecordGetInputName(c))
                If (BASS_RecordGetInput(c, ByVal 0) And BASS_INPUT_OFF) = 0 Then
                    cmbInput.ListIndex = c  ' this 1 is currently "on"
                    input_ = c
                    Call UpdateInputInfo    ' display info
                End If
                c = c + 1
            Wend
        End If
    Else
        MsgBox ("Não existem dispositivos disponíveis")
    End If
End Sub

Private Sub btnRecord_Click()
   If (Not diretorio Is Nothing) Then
        If (rchan = 0) Then
             Call StartRecording(diretorio, ComboQualidade.ItemData(ComboQualidade.ListIndex))
             TimerGravacao.Enabled = True
             TimerSegundos.Enabled = True
             'desabilita ajustes
             ComboDivisao.Enabled = False
             ComboQualidade.Enabled = False
             cmbInput.Enabled = False
             Contador = 0
         Else
             Call StopRecording
             TimerGravacao.Enabled = False
             TimerSegundos.Enabled = False
             CheckProgramado.value = 0
             ''habilita ajustes
             ComboDivisao.Enabled = True
             ComboQualidade.Enabled = True
             cmbInput.Enabled = True
             'reseta medidores
             ProgressBarLeft.value = 0.0001
             ProgressBarRight.value = 0.0001
             Labelhora = "00:00:00"
         End If
    Else
        MsgBox ("O diretório de trabalho ainda não foi definido")
    End If
End Sub

'atualiza a label q ue diz qual input usado
Private Sub cmbInput_Click()
If (input_ > -1) Then
        input_ = cmbInput.ListIndex
        ' enable the selected input
        Dim i As Integer
        For i = 0 To cmbInput.ListCount - 1
            Call BASS_RecordSetInput(i, BASS_INPUT_OFF, -1) ' 1st disable all inputs, then...
        Next i
        Call BASS_RecordSetInput(input_, BASS_INPUT_ON, -1) ' enable the selected input
        Call UpdateInputInfo
    End If
End Sub

'sai do programa
Private Sub mnuSair_Click()
    Unload Me
End Sub

'mostra o form de informações
Private Sub mnusobre_Click()
    Sobre.Show 0, Main
End Sub

' set input source level
Private Sub sldInputLevel_Scroll()
    Call BASS_RecordSetInput(input_, 0, sldInputLevel.value / 100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' release all BASS stuff
    If (BASS_ChannelIsActive(rchan) = BASS_ACTIVE_PLAYING) Then
        Call StopRecording
    End If
    Call BASS_RecordFree
    Call BASS_Free
End Sub
'recarregando os áudios quando vai pra segunda tab
Private Sub SSTab1_Click(PreviousTab As Integer)
    If (SSTab1.Tab = 1) Then
        If (Not diretorio Is Nothing) Then
            Call recuperaArquivos(TextDiretorio.Text)
        End If
    End If
End Sub
' timer dos eventos programados
Private Sub TimerEventos_Timer()
    Dim agora As Date
    agora = CDate(Hour(Now) & ":" & Minute(Now) & ":" & Second(Now))
    If (agora >= CDate(TextIni) And agora <= CDate(TextFim)) Then
        If (BASS_ChannelIsActive(rchan) = BASS_ACTIVE_STOPPED Or rchan = 0) Then
            Call StartRecording(diretorio, ComboQualidade.ItemData(ComboQualidade.ListIndex))
            TimerGravacao.Enabled = True
            TimerSegundos.Enabled = True
            'desabilita ajustes
            ComboDivisao.Enabled = False
            ComboQualidade.Enabled = False
            cmbInput.Enabled = False
        End If
    'se não está dentro do horário programado
    Else
        'se se tratar do evento que está atualmente executando
        If (BASS_ChannelIsActive(rchan) = BASS_ACTIVE_PLAYING) Then
            Call StopRecording
            TimerGravacao.Enabled = False
            TimerSegundos.Enabled = False
            ''habilita ajustes
            ComboDivisao.Enabled = True
            ComboQualidade.Enabled = True
            cmbInput.Enabled = True
            'reseta medidores
            ProgressBarLeft.value = 0.0001
            ProgressBarRight.value = 0.0001
            Labelhora = "00:00:00"
        End If
    End If
End Sub

'timer utilizado para eventos dinâmicos durante a gravação
Private Sub TimerGravacao_Timer()
    If (LoWord(getLevel(rchan)) < 32767) Then
        ProgressBarLeft.value = LoWord(getLevel(rchan))
        LabelClipL.BackColor = ColorConstants.vbBlue
    Else
        LabelClipL.BackColor = ColorConstants.vbRed
    End If
    
    If (HiWord(getLevel(rchan)) < 32767) Then
        ProgressBarRight.value = HiWord(getLevel(rchan))
        LabelClipR.BackColor = ColorConstants.vbBlue
    Else
        LabelClipR.BackColor = ColorConstants.vbRed
    End If
End Sub

Private Sub TimerRelogio_Timer()
    LabelRelogio.Caption = Format(Now, "Long Time")
End Sub

'timer utilizado para eventos informações durante a gravação
Private Sub TimerSegundos_Timer()
    LabelInfos.Caption = " Letra associada: " & disco.DriveLetter _
    & vbCrLf & " Nome do disco: " & disco.VolumeName _
    & vbCrLf & " Espaço livre: " & FormatNumber(disco.AvailableSpace / 1024, 0) _
    & vbCrLf & " Espaço Total: " & FormatNumber(disco.TotalSize / 1024, 0)
    
    If (ComboDivisao.ItemData(ComboDivisao.ListIndex) > 0) Then
        Contador = Contador + 1
        Labelhora.Caption = formatHora(Contador)
        'fazendo a quebra do arquivo de acordo com a divisão escolhida
        If (Contador >= ComboDivisao.ItemData(ComboDivisao.ListIndex)) Then
            'salvando e abrindo outra gravação
            Contador = 0
            TimerEventos.Enabled = False
            TimerGravacao.Enabled = False
            Call StopRecording
            Call StartRecording(diretorio, ComboQualidade.ItemData(ComboQualidade.ListIndex))
            TimerGravacao.Enabled = True
            If (CheckProgramado.value = 1) Then Enabled = True 'habilita somente se agendamento ativado
        End If
    End If
End Sub

'Recupera arquivos de uma pasta
Private Sub recuperaArquivos(nameDiret As String)
    Dim arquivo As Scripting.file
    Set disco = myFSO.GetDrive(Left$(nameDiret, 3))
    Set diretorio = myFSO.GetFolder(nameDiret)
    
    'listando arquivos
    ListViewAudios.ListItems.Clear
    For Each arquivo In diretorio.Files
        If (Right$(UCase$(arquivo.name), 3) Like "MP3") Then
            ListViewAudios.ListItems.Add , arquivo.Path, arquivo.name
        End If
    Next
End Sub

'Formata hora
Private Function formatHora(segundos As Integer) As String
    Dim hor, min, seg As Integer
    hor = segundos \ 3600
    min = (segundos Mod 3600) \ 60
    seg = (segundos Mod 3600) Mod 60
    formatHora = Format$(hor, "00") & ":" & Format$(min, "00") & ":" & Format$(seg, "00")
End Function

Private Sub reloadDeviceInputs()
    If (ComboDevices.ListCount > 0) Then
        'selecionando dispositivo de gravação
        If (BASS_RecordInit(ComboDevices.ListIndex) = 0) Or (BASS_Init(ComboDevices.ListIndex, 44100, 0, Me.hWnd, 0) = 0) Then
            Call Error_("O dispositivo não pode ser iniciado")
            End
        Else
            'carregando uma lista de entradas
            Dim c As Integer
            input_ = -1
            While BASS_RecordGetInputName(c)
                cmbInput.AddItem VBStrFromAnsiPtr(BASS_RecordGetInputName(c))
                If (BASS_RecordGetInput(c, ByVal 0) And BASS_INPUT_OFF) = 0 Then
                    cmbInput.ListIndex = c  ' this 1 is currently "on"
                    input_ = c
                    Call UpdateInputInfo    ' display info
                End If
                c = c + 1
            Wend
        End If
    Else
        MsgBox ("Não existem dispositivos disponíveis")
    End If
End Sub
