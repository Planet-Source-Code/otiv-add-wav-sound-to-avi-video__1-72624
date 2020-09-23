Attribute VB_Name = "mAddWav"
Option Explicit


Private Declare Sub AVIFileInit Lib "avifil32.dll" ()


Private Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long


Private Declare Function AVIFileCreateStream Lib "avifil32.dll" Alias "AVIFileCreateStreamA" _
    (ByVal pfile As Long, ByRef ppavi As Long, ByRef psi As AVI_STREAM_INFO) As Long


Private Declare Function AVIStreamSetFormat Lib "avifil32.dll" (ByVal pavi As Long, _
    ByVal lPos As Long, _
    ByRef lpFormat As Any, _
    ByVal cbFormat As Long) As Long


Private Declare Function AVIStreamWrite Lib "avifil32.dll" (ByVal pavi As Long, _
    ByVal lStart As Long, _
    ByVal lSamples As Long, _
    ByVal lpBuffer As Long, _
    ByVal cbBuffer As Long, _
    ByVal dwFlags As Long, _
    ByRef plSampWritten As Long, _
    ByRef plBytesWritten As Long) As Long


Private Declare Function AVIStreamReadFormat Lib "avifil32.dll" (ByVal pAVIStream As Long, _
    ByVal lPos As Long, _
    ByRef lpFormat As PCMWAVEFORMAT, _
    ByRef cbFormat As Long) As Long


Private Declare Function AVIStreamRead Lib "avifil32.dll" (ByVal pAVIStream As Long, _
    ByVal lStart As Long, _
    ByVal lSamples As Long, _
    ByVal lpBuffer As Long, _
    ByVal cbBuffer As Long, _
    ByRef pBytesWritten As Long, _
    ByRef pSamplesWritten As Long) As Long


Private Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long


Private Declare Function AVIStreamInfo Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef psi As AVI_STREAM_INFO, ByVal lSize As Long) As Long


Private Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long


Private Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long


Private Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pfile As Long) As Long


Private Declare Sub AVIFileExit Lib "avifil32.dll" ()


Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long


Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long


Private Type WAVEFORMAT
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    End Type


Private Type PCMWAVEFORMAT
    wf As WAVEFORMAT
    wBitsPerSample As Integer
    End Type


Private Type AVI_RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type


Private Type AVI_STREAM_INFO
    fccType As Long
    fccHandler As Long
    dwFlags As Long
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As AVI_RECT
    dwEditCount As Long
    dwFormatChangeCount As Long
    szName As String * 64
    End Type
    Private Const AVIERR_OK As Long = 0&
    Private Const OF_READWRITE As Long = &H2
    Private Const AVIIF_KEYFRAME As Long = &H10
    Private Const streamtypeVIDEO As Long = 1935960438
    Private Const streamtypeAUDIO As Long = 1935963489
    Private Const streamtypeMIDI As Long = 1935960429
    Private Const streamtypeTEXT As Long = 1937012852
    Private Const GMEM_FIXED = &H0
    Private Const GMEM_ZEROINIT = &H40
    Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
    'To use this function, need 2 files
    'existing avi video file(without sound)
    'existing wav audio file


Public Sub AddAudioStream(ByVal AVIFilePath As String, ByVal WAVFilePath As String) 'AddAudioStream(Destination AVI FilePath, Source WAV FilePath)
    On Error GoTo errHandler
    Dim StreamInfo As AVI_STREAM_INFO
    Dim StreamFormat As PCMWAVEFORMAT
    Dim StreamLength As Long
    Dim WaveData As Long
    Dim AVIFile As Long
    Dim AudioFile As Long
    Dim AudioStream As Long
    Dim AVIStream As Long
    Call AVIFileInit 'Initialize the AVI library.


    If AVIFileOpen(AVIFile, AVIFilePath, OF_READWRITE, ByVal 0&) = AVIERR_OK Then 'open the avi file


        If AVIFileOpen(AudioFile, WAVFilePath, OF_READWRITE, ByVal 0&) = AVIERR_OK Then 'open the wave file
            If AVIFileGetStream(AudioFile, AudioStream, streamtypeAUDIO, 0) <> AVIERR_OK Then GoTo errHandler 'get the audio stream
            If AVIStreamInfo(AudioStream, StreamInfo, Len(StreamInfo)) <> AVIERR_OK Then GoTo errHandler 'read the stream's header information
            AVIStreamReadFormat AudioStream, 0, StreamFormat, Len(StreamFormat) 'read the stream's format data
            StreamLength = AVIStreamLength(AudioStream) * StreamInfo.dwSampleSize 'get the length of the stream
            WaveData = GlobalAlloc(GPTR, StreamLength) 'get pointer To the wave data
            AVIStreamRead AudioStream, 0, StreamLength, WaveData, StreamLength, 0, 0 'read audio data from the stream
            AVIFileCreateStream AVIFile, AVIStream, StreamInfo 'create new stream
            AVIStreamSetFormat AVIStream, 0, StreamFormat, Len(StreamFormat) 'set the format of new stream
            AVIStreamWrite AVIStream, 0, StreamLength, WaveData, StreamLength, AVIIF_KEYFRAME, 0, 0 'copy the raw wave data To new stream
            GlobalFree WaveData 'release the wave data pointer
            AVIStreamRelease (AudioStream)
            AVIStreamRelease (AVIStream)
            AVIFileRelease AudioFile
            AVIFileRelease AVIFile
            Call AVIFileExit 'close the AVI library
        End If
    End If
errHandler:

End Sub

