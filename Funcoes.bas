Attribute VB_Name = "Funcoes"
Option Explicit
Public input_ As Long
Public command As String    ' oggenc (OGG), lame (MP3)
Public file As String       ' OGG,MP3,ACM
Public rchan As Long        ' recording/encoding channel
Public chan As Long         ' playback channel

' display error messages
Public Sub Error_(ByVal es As String)
    Call MsgBox(es & vbCrLf & vbCrLf & "error code: " & BASS_ErrorGetCode, vbExclamation, "Error")
End Sub

'funcão callback passada para a função do bass que grava
Public Function RecordingCallback(ByVal channel As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
    RecordingCallback = BASS_Encode_IsActive(channel)   ' continue recording if encoder is alive
End Function

'CAPTURA O NÍVEL DO VOLUME E MOSTRA O DISPOSITIVO ESCOLHIDO NA LABEL
Public Sub UpdateInputInfo()
    Dim it As Long
    Dim level As Single
    it = BASS_RecordGetInput(input_, level)
    Main.sldInputLevel.value = level * 100 ' set the level slider
    
    Dim type_ As String
    Select Case (it And BASS_INPUT_TYPE_MASK)
        Case BASS_INPUT_TYPE_DIGITAL:
            type_ = "digital"
        Case BASS_INPUT_TYPE_LINE:
            type_ = "Entrada de Linha"
        Case BASS_INPUT_TYPE_MIC:
            type_ = "Entrada de microfone"
        Case BASS_INPUT_TYPE_SYNTH:
            type_ = "Sintetizador Midi"
        Case BASS_INPUT_TYPE_CD:
            type_ = "Gravação do CD"
        Case BASS_INPUT_TYPE_PHONE:
            type_ = "Telefone"
        Case BASS_INPUT_TYPE_SPEAKER:
            type_ = "pc speaker"
        Case BASS_INPUT_TYPE_WAVE:
            type_ = "wave/pcm"
        Case BASS_INPUT_TYPE_AUX:
            type_ = "Entrada Auxiliar"
        Case BASS_INPUT_TYPE_ANALOG:
            type_ = "Mixagem"
        Case Else:
            type_ = "undefined"
    End Select
    Main.lblInputType.Caption = type_ ' display the type
End Sub
'captura o nível de gravação
Public Function getLevel(channel) As Long
    If (rchan <> 0) Then
        getLevel = BASS_ChannelGetLevel(channel)
    End If
End Function

' start recording
Public Sub StartRecording(diret As Scripting.Folder, bitRate As Integer)
    'definindo linha de comando para gravação
    file = diret.ShortPath & "\" & Format(Now, "ddmmyyyy-hhmmss") & ".mp3" 'nome do arquivo
    command = "lame.exe --alt-preset cbr " & bitRate & " - " & file ' linha de comando do encoder
    
    ' free old recording
    If (chan) Then
        Call BASS_StreamFree(chan)
        chan = 0
    End If
    
    ' Start recording @ 44100hz 16-bit stereo (paused to add encoder first)
    rchan = BASS_RecordStart(44100, 2, BASS_RECORD_PAUSE, AddressOf RecordingCallback, 0)
    
    If (rchan = 0) Then
        Call Error_("Couldn't start recording")
        Exit Sub
    End If

    'iniciando a gravação em mp3
    If (BASS_Encode_Start(rchan, command, BASS_ENCODE_AUTOFREE, 0, 0) = 0) Then ' start the OGG/MP3 encoder
        Call Error_("A gravação não pôde ser iniciada pois não foi encontrado o aplicativo lame.exe")
        Call BASS_ChannelStop(rchan)
        rchan = 0
        Exit Sub
    End If

    ' resume recoding
    Call BASS_ChannelPlay(rchan, BASSFALSE)
End Sub

' stop recording
Public Sub StopRecording()
    ' stop recording & encoding
    Call BASS_ChannelStop(rchan)
    rchan = 0

    ' create a stream from the recording
    chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(file), 0, 0, 0)
End Sub
