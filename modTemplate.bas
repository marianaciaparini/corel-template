Attribute VB_Name = "modTemplate"
' Corel Template - CorelDRAW 2018+ (tested on 20.1)
' Page 1 is the template and is never changed.
' For each chosen image: duplicates page 1 at the end of the document, imports the image
' and places it inside the frame named FRAME_NAME (PowerClip), using the fit mode chosen
' in the dialog: Fill, Fit, Same as example, Manual.
' Coordinates match CorelDRAW's property bar: millimetres, page origin, object centre.
Option Explicit

' ===== Settings =====
' Language of the dialog and messages: "en" (English) or "pt" (Portugues)
Public Const LANG As String = "en"
' Name of the frame object on page 1 (set it in the Object Manager)
Public Const FRAME_NAME As String = "FOTO"

Private Const APP As String = "CorelTemplate"
Private Const SEC As String = "Caixa"

#If VBA7 Then
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#Else
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#End If

' ===== Main macro =====
Public Sub NovaCopiaComImagem()
    Dim doc As Document, pgModelo As Page, foto As Shape, ex As Shape
    Dim arquivos As Collection, i As Long
    Dim unidadeAntes As cdrUnit, refAntes As cdrReferencePoint
    Dim pg As Page, moldura As Shape, img As Shape, primeiraPg As Page
    Dim frm As frmTemplate
    Dim modo As String, ancH As Double, ancV As Double
    Dim vX As Double, vY As Double, vW As Double, vH As Double
    Dim travar As Boolean, mestre As String
    Dim w As Double, h As Double, s As Double, grupoAberto As Boolean
    Dim fx As Double, fy As Double, fw As Double, fh As Double

    If Documents.Count = 0 Then MsgBox T("Abra o documento do modelo primeiro.", "Open the template document first."), vbExclamation: Exit Sub
    Set doc = ActiveDocument
    Set pgModelo = doc.Pages(1)
    Set foto = pgModelo.Shapes.FindShape(Name:=FRAME_NAME)
    If foto Is Nothing Then
        MsgBox T("Nao achei um objeto chamado """ & FRAME_NAME & """ na pagina 1." & vbCr & _
                 "De esse nome a moldura no Gerenciador de objetos.", _
                 "No object named """ & FRAME_NAME & """ on page 1." & vbCr & _
                 "Give the frame that name in the Object Manager."), vbExclamation
        Exit Sub
    End If

    Set arquivos = EscolherImagens()
    If arquivos.Count = 0 Then Exit Sub

    unidadeAntes = doc.Unit
    refAntes = doc.ReferencePoint
    On Error GoTo Falha
    doc.Unit = cdrMillimeter
    doc.ReferencePoint = cdrCenter
    doc.BeginCommandGroup T("Nova copia com imagem", "New copy with image")
    grupoAberto = True

    For i = 1 To arquivos.Count
        Set pg = CopiarPagina(doc, pgModelo)
        If i = 1 Then Set primeiraPg = pg
        Set moldura = pg.Shapes.FindShape(Name:=FRAME_NAME)
        If moldura Is Nothing Then Err.Raise vbObjectError + 1, , T("A copia da pagina nao tem a moldura ", "The page copy has no frame ") & FRAME_NAME & "."
        LimparMoldura moldura

        pg.Activate
        moldura.Layer.Activate
        moldura.Layer.Import arquivos(i)
        Set img = ActiveShape
        If img Is Nothing Then Err.Raise vbObjectError + 2, , T("Falhou a importacao de ", "Import failed: ") & arquivos(i)

        If i = 1 Then
            ' Dialog shows once; the 1st image is the aspect-ratio reference
            Set ex = ImagemExemplo(foto)
            Set frm = New frmTemplate
            If ex Is Nothing Then
                frm.Preparar GetSetting(APP, SEC, "Modo", "FILL"), False, 0, 0, 0, 0, _
                    LerNum("X", foto.CenterX), LerNum("Y", foto.CenterY), _
                    LerNum("W", img.SizeWidth), LerNum("H", img.SizeHeight), _
                    (GetSetting(APP, SEC, "Travar", "1") = "1"), img.SizeHeight / img.SizeWidth, _
                    GetSetting(APP, SEC, "Mestre", "W"), _
                    CLng(GetSetting(APP, SEC, "AncH", "1")), CLng(GetSetting(APP, SEC, "AncV", "2")), _
                    InfoTexto(arquivos, img)
            Else
                frm.Preparar GetSetting(APP, SEC, "Modo", "FILL"), True, ex.CenterX, ex.CenterY, ex.SizeWidth, ex.SizeHeight, _
                    LerNum("X", foto.CenterX), LerNum("Y", foto.CenterY), _
                    LerNum("W", img.SizeWidth), LerNum("H", img.SizeHeight), _
                    (GetSetting(APP, SEC, "Travar", "1") = "1"), img.SizeHeight / img.SizeWidth, _
                    GetSetting(APP, SEC, "Mestre", "W"), _
                    CLng(GetSetting(APP, SEC, "AncH", "1")), CLng(GetSetting(APP, SEC, "AncV", "2")), _
                    InfoTexto(arquivos, img)
            End If
            frm.Show vbModal
            If Not frm.OK Then
                Unload frm
                primeiraPg.Delete
                GoTo Fim
            End If
            modo = frm.Modo
            vX = frm.ValX: vY = frm.ValY: vW = frm.ValW: vH = frm.ValH
            travar = frm.Travado: mestre = frm.Mestre
            ancH = FracaoH(frm.AncoraH): ancV = FracaoV(frm.AncoraV)
            SaveSetting APP, SEC, "Modo", modo
            SaveSetting APP, SEC, "AncH", CStr(frm.AncoraH)
            SaveSetting APP, SEC, "AncV", CStr(frm.AncoraV)
            SaveSetting APP, SEC, "Travar", IIf(travar, "1", "0")
            SaveSetting APP, SEC, "Mestre", mestre
            If modo = "MANUAL" Then
                SaveSetting APP, SEC, "X", CStr(vX)
                SaveSetting APP, SEC, "Y", CStr(vY)
                SaveSetting APP, SEC, "W", CStr(vW)
                SaveSetting APP, SEC, "H", CStr(vH)
            End If
            Unload frm
        End If

        If modo = "FILL" Or modo = "FIT" Then
            ' Frame bounding box (for round frames: the square around it)
            fx = moldura.LeftX: fy = moldura.BottomY
            fw = moldura.SizeWidth: fh = moldura.SizeHeight
            If modo = "FILL" Then
                s = Maior(fw / img.SizeWidth, fh / img.SizeHeight)   ' cover, crops the excess
            Else
                s = Menor(fw / img.SizeWidth, fh / img.SizeHeight)   ' whole image, leaves margins
            End If
            w = img.SizeWidth * s: h = img.SizeHeight * s
            img.SetSizeEx img.CenterX, img.CenterY, w, h
            ' Anchor: 0 = left/bottom, 0.5 = centre, 1 = right/top
            img.SetPositionEx cdrBottomLeft, fx + (fw - w) * ancH, fy + (fh - h) * ancV
        Else
            ' EXAMPLE / MANUAL: locked ratio -> last edited field rules, the other follows THIS image
            If travar Then
                If mestre = "H" Then
                    h = vH: w = vH * img.SizeWidth / img.SizeHeight
                Else
                    w = vW: h = vW * img.SizeHeight / img.SizeWidth
                End If
            Else
                w = vW: h = vH
            End If
            img.SetSizeEx img.CenterX, img.CenterY, w, h
            img.SetPositionEx cdrCenter, vX, vY
        End If
        img.AddToPowerClip moldura, cdrFalse
    Next i

Fim:
    If grupoAberto Then doc.EndCommandGroup
    doc.Unit = unidadeAntes
    doc.ReferencePoint = refAntes
    ActiveWindow.Refresh
    Exit Sub

Falha:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    If grupoAberto Then doc.EndCommandGroup
    doc.Unit = unidadeAntes
    doc.ReferencePoint = refAntes
    MsgBox T("Parou com erro: ", "Stopped with error: ") & msg & vbCr & _
           T("Ctrl+Z desfaz o que a macro fez.", "Ctrl+Z undoes what the macro did."), vbCritical
End Sub

' ===== Duplicates the template page at the end, layer by layer, without touching the original =====
Private Function CopiarPagina(doc As Document, pgModelo As Page) As Page
    Dim novo As Page, lr As Layer, dst As Layer, j As Long
    Set novo = doc.AddPages(1)
    novo.SetSize pgModelo.SizeWidth, pgModelo.SizeHeight
    For Each lr In pgModelo.Layers
        If Not lr.IsSpecialLayer Then
            Set dst = AcharCamada(novo, lr.Name)
            If dst Is Nothing Then Set dst = novo.CreateLayer(lr.Name)
            ' bottom to top, keeps the stacking order
            For j = lr.Shapes.Count To 1 Step -1
                lr.Shapes(j).CopyToLayer dst
            Next j
        End If
    Next lr
    Set CopiarPagina = novo
End Function

Private Function AcharCamada(pg As Page, nome As String) As Layer
    Dim lr As Layer
    For Each lr In pg.Layers
        If lr.Name = nome Then Set AcharCamada = lr: Exit Function
    Next lr
End Function

' Example image inside the frame on page 1 (read only; the template is not changed)
Private Function ImagemExemplo(foto As Shape) As Shape
    On Error Resume Next
    If foto.PowerClip Is Nothing Then Exit Function
    If foto.PowerClip.Shapes.Count = 0 Then Exit Function
    Set ImagemExemplo = foto.PowerClip.Shapes(1)
End Function

' If the template has an example image inside the frame, the copy comes out without it
Private Sub LimparMoldura(moldura As Shape)
    On Error Resume Next
    If Not moldura.PowerClip Is Nothing Then
        If moldura.PowerClip.Shapes.Count > 0 Then moldura.PowerClip.Shapes.All.Delete
    End If
End Sub

' ===== Windows file dialog with multi-select =====
Private Function EscolherImagens() As Collection
    Dim ofn As OPENFILENAME, r As Long, s As String, partes() As String
    Dim res As New Collection, pasta As String, k As Long
    ofn.lStructSize = LenB(ofn)
    ofn.lpstrFilter = T("Imagens", "Images") & vbNullChar & "*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.bmp;*.gif" & vbNullChar & _
                      T("Todos os arquivos", "All files") & vbNullChar & "*.*" & vbNullChar & vbNullChar
    ofn.nFilterIndex = 1
    ofn.lpstrFile = String$(65536, vbNullChar)
    ofn.nMaxFile = Len(ofn.lpstrFile)
    ofn.lpstrTitle = T("Escolha as imagens (Ctrl ou Shift para varias)", "Choose the images (Ctrl or Shift for several)")
    ofn.lpstrInitialDir = GetSetting(APP, SEC, "Pasta", "")
    ofn.flags = &H80000 Or &H200 Or &H1000 Or &H4   ' EXPLORER, ALLOWMULTISELECT, FILEMUSTEXIST, HIDEREADONLY
    r = GetOpenFileName(ofn)
    Set EscolherImagens = res
    If r = 0 Then Exit Function

    s = ofn.lpstrFile
    k = InStr(s, vbNullChar & vbNullChar)
    If k > 0 Then s = Left$(s, k - 1)
    partes = Split(s, vbNullChar)
    If UBound(partes) = 0 Then
        res.Add partes(0)
        pasta = Left$(partes(0), InStrRev(partes(0), "\") - 1)
    Else
        pasta = partes(0)
        For k = 1 To UBound(partes)
            If Len(partes(k)) > 0 Then res.Add pasta & "\" & partes(k)
        Next k
    End If
    SaveSetting APP, SEC, "Pasta", pasta
End Function

' ===== Helpers =====
Public Function T(pt As String, en As String) As String
    If LANG = "pt" Then T = pt Else T = en
End Function

' Horizontal anchor index -> fraction: 0 left, 1 centre, 2 right
Private Function FracaoH(idx As Long) As Double
    Select Case idx
        Case 0: FracaoH = 0
        Case 2: FracaoH = 1
        Case Else: FracaoH = 0.5
    End Select
End Function

' Vertical anchor index -> fraction (Y grows upwards): 0 top, 1 upper third, 2 centre, 3 bottom
Private Function FracaoV(idx As Long) As Double
    Select Case idx
        Case 0: FracaoV = 1
        Case 1: FracaoV = 2 / 3
        Case 3: FracaoV = 0
        Case Else: FracaoV = 0.5
    End Select
End Function

Private Function InfoTexto(arquivos As Collection, img As Shape) As String
    InfoTexto = T("Referencia: ", "Reference: ") & NomeArquivo(arquivos(1)) & _
                " (" & arquivos.Count & T(" imagem(ns))", " image(s))") & vbCr & _
                T("Original: ", "Original: ") & Fmt(img.SizeWidth) & " x " & Fmt(img.SizeHeight) & " mm"
End Function

Private Function LerNum(chave As String, padrao As Double) As Double
    Dim s As String
    s = GetSetting(APP, SEC, chave, "")
    If s = "" Then LerNum = padrao Else LerNum = Val(Replace(s, ",", "."))
End Function

Public Function Fmt(v As Double) As String
    If LANG = "pt" Then
        Fmt = Replace(Format$(v, "0.000"), ".", ",")
    Else
        Fmt = Replace(Format$(v, "0.000"), ",", ".")
    End If
End Function

Private Function NomeArquivo(caminho As String) As String
    NomeArquivo = Mid$(caminho, InStrRev(caminho, "\") + 1)
End Function

Private Function Maior(a As Double, b As Double) As Double
    If a > b Then Maior = a Else Maior = b
End Function

Private Function Menor(a As Double, b As Double) As Double
    If a < b Then Menor = a Else Menor = b
End Function
