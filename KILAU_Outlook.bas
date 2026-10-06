Attribute VB_Name = "KILAU_Outlook"
' =====================================================================
'  KILAU (Kerja Instan, Laporan Akurat & Utuh)
'  Macro Outlook untuk program cashback Emas30/Emas50
'  BBG - Bank Syariah Indonesia
'
'  1) SimpanLampiranEmas3050
'     Menyimpan lampiran Data_Eligible_* dan File_Multiposting_* dari
'     email DAS ke  Documents\Emas30_50
'
'  2) BuatBalasanMultiposting
'     Untuk setiap email "[DAS] - FILE MULTIPOSTING ..." membuat balasan
'     (Reply All) yang berisi:
'       - file multiposting FIX hari yang sama  (dari folder Balasan)
'       - berita acara                          (PDF di folder BA)
'       - isi email per hari yang disusun tools KILAU
'     Default: balasan disimpan sebagai DRAFT untuk dicek dulu.
'
'  3) UjiCobaBalasan
'     Sama seperti nomor 2, tetapi maksimal 3 email, penerima diganti
'     menjadi email lo sendiri, subjek diberi [UJI COBA], dan email DAS
'     aslinya tidak ditandai. Aman untuk percobaan.
'
'  Struktur folder (dibuat otomatis):
'     Documents\Emas30_50\           lampiran asli dari DAS
'     Documents\Emas30_50\Balasan\   isi ZIP "paket balasan" dari tools
'     Documents\Emas30_50\BA\        file PDF berita acara yang berlaku
'
'  Syarat: Outlook versi klasik (desktop).
' =====================================================================
Option Explicit

Private Const FOLDER_DASAR As String = ""      ' kosong = Documents\Emas30_50
Private Const HARI_DEFAULT As Long = 90
Private Const KIRIM_LANGSUNG As Boolean = False ' True = langsung kirim, False = simpan draft
Private Const KATEGORI As String = "KILAU Dibalas"

' ---------------------------------------------------------------------
Private Function FolderDasar() As String
    Dim f As String
    f = FOLDER_DASAR
    If f = "" Then f = Environ("USERPROFILE") & "\Documents\Emas30_50"
    If Dir(f, vbDirectory) = "" Then MkDir f
    If Dir(f & "\Balasan", vbDirectory) = "" Then MkDir f & "\Balasan"
    If Dir(f & "\BA", vbDirectory) = "" Then MkDir f & "\BA"
    FolderDasar = f
End Function

Private Function TanyaHari() As Long
    Dim j As String
    j = InputBox("Periksa email berapa hari ke belakang?", "KILAU", CStr(HARI_DEFAULT))
    If j = "" Or Not IsNumeric(j) Then
        TanyaHari = -1
    Else
        TanyaHari = CLng(j)
    End If
End Function

' =====================================================================
'  1) SIMPAN LAMPIRAN
' =====================================================================
Public Sub SimpanLampiranEmas3050()
    Dim fld As Outlook.MAPIFolder, itms As Outlook.Items, itm As Object
    Dim att As Outlook.Attachment, tujuan As String, nama As String, jalur As String
    Dim hari As Long, disimpan As Long, dilewati As Long, emailCocok As Long, ada As Boolean

    MsgBox "Pilih folder Outlook tempat email laporan Emas30/Emas50 dari DAS berada.", vbInformation, "KILAU"
    Set fld = Application.GetNamespace("MAPI").PickFolder
    If fld Is Nothing Then Exit Sub
    hari = TanyaHari(): If hari < 0 Then Exit Sub
    tujuan = FolderDasar()

    Set itms = fld.Items
    itms.Sort "[ReceivedTime]", True
    For Each itm In itms
        If TypeOf itm Is Outlook.MailItem Then
            If itm.ReceivedTime < Date - hari Then Exit For
            ada = False
            For Each att In itm.Attachments
                nama = att.FileName
                If (LCase(nama) Like "data_eligible_*.xls*") Or (LCase(nama) Like "file_multiposting_*.xls*") Then
                    If InStr(1, nama, "(FIX)", vbTextCompare) = 0 Then
                        ada = True
                        jalur = tujuan & "\" & nama
                        If Dir(jalur) = "" Then
                            att.SaveAsFile jalur: disimpan = disimpan + 1
                        Else
                            dilewati = dilewati + 1
                        End If
                    End If
                End If
            Next att
            If ada Then emailCocok = emailCocok + 1
        End If
    Next itm

    MsgBox "Selesai." & vbCrLf & vbCrLf & "Email berisi laporan: " & emailCocok & vbCrLf & _
           "File baru disimpan: " & disimpan & vbCrLf & "File sudah ada (dilewati): " & dilewati & vbCrLf & vbCrLf & _
           "Lokasi: " & tujuan & vbCrLf & vbCrLf & "Buka tools KILAU, klik ""Pilih folder"", lalu pilih folder ini.", vbInformation, "KILAU"
    Shell "explorer.exe """ & tujuan & """", vbNormalFocus
End Sub

' =====================================================================
'  2) BUAT BALASAN PER HARI
' =====================================================================
Public Sub BuatBalasanMultiposting()
    ProsesBalasan False, "", 0
End Sub

Public Sub UjiCobaBalasan()
    Dim alamat As String
    On Error Resume Next
    alamat = Application.Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress
    On Error GoTo 0
    alamat = InputBox("Mode UJI COBA: maksimal 3 balasan dikirim HANYA ke alamat ini (bukan ke DAS/DBO)." & vbCrLf & vbCrLf & _
                      "Alamat email penerima uji coba:", "KILAU - Uji coba", alamat)
    If Trim(alamat) = "" Then Exit Sub
    ProsesBalasan True, Trim(alamat), 3
End Sub

Private Sub ProsesBalasan(ByVal uji As Boolean, ByVal alamatUji As String, ByVal batas As Long)
    Dim dasar As String, folderBalasan As String, baPath As String, baNama As String
    Dim fld As Outlook.MAPIFolder, itms As Outlook.Items, itm As Object, rep As Outlook.MailItem
    Dim subj As String, prog As String, tgl As String, fixPath As String, txtPath As String
    Dim isi As String, meta As String, teks As String, nDas As Long, nAsli As Long
    Dim hari As Long, dibuat As Long, sudah As Long, nol As Long
    Dim belumFix As String, selisih As String

    dasar = FolderDasar()
    folderBalasan = dasar & "\Balasan"
    baNama = Dir(dasar & "\BA\*.pdf")
    If baNama = "" Then
        MsgBox "File berita acara (PDF) belum ada di:" & vbCrLf & dasar & "\BA" & vbCrLf & vbCrLf & _
               "Simpan PDF berita acara yang berlaku di folder itu, lalu jalankan lagi.", vbExclamation, "KILAU"
        Shell "explorer.exe """ & dasar & "\BA""", vbNormalFocus
        Exit Sub
    End If
    baPath = dasar & "\BA\" & baNama
    If Dir(folderBalasan & "\*(FIX).xlsx") = "" Then
        MsgBox "Folder Balasan masih kosong:" & vbCrLf & folderBalasan & vbCrLf & vbCrLf & _
               "Unduh ""paket balasan per hari (.zip)"" dari tools KILAU, lalu ekstrak isinya ke folder itu.", vbExclamation, "KILAU"
        Shell "explorer.exe """ & folderBalasan & """", vbNormalFocus
        Exit Sub
    End If

    MsgBox "Pilih folder Outlook tempat email ""[DAS] - FILE MULTIPOSTING"" berada.", vbInformation, "KILAU"
    Set fld = Application.GetNamespace("MAPI").PickFolder
    If fld Is Nothing Then Exit Sub
    hari = TanyaHari(): If hari < 0 Then Exit Sub

    Set itms = fld.Items
    itms.Sort "[ReceivedTime]", True
    For Each itm In itms
        If TypeOf itm Is Outlook.MailItem Then
            If itm.ReceivedTime < Date - hari Then Exit For
            If batas > 0 And dibuat >= batas Then Exit For
            subj = itm.Subject
            If UCase(Left(Trim(subj), 5)) = "[DAS]" And InStr(1, subj, "FILE MULTIPOSTING", vbTextCompare) > 0 Then
                If Not uji And InStr(1, itm.Categories, KATEGORI, vbTextCompare) > 0 Then
                    sudah = sudah + 1
                Else
                    prog = ""
                    If InStr(1, subj, "EMAS50", vbTextCompare) > 0 Then
                        prog = "EMAS50"
                    ElseIf InStr(1, subj, "EMAS30", vbTextCompare) > 0 Then
                        prog = "EMAS30"
                    End If
                    tgl = AmbilTanggal(subj)
                    fixPath = CariFix(folderBalasan, prog, tgl)
                    If AmbilAngka(itm.Body, "eligible sebanyak") = 0 Then
                        nol = nol + 1
                    ElseIf prog = "" Or tgl = "" Or fixPath = "" Then
                        belumFix = belumFix & "- " & prog & " " & tgl & vbCrLf
                    Else
                        txtPath = Left(fixPath, Len(fixPath) - 5) & ".txt"
                        isi = BacaUtf8(txtPath)
                        meta = "": teks = isi
                        If Left(isi, 6) = "#KILAU" Then
                            meta = Split(isi, vbCrLf)(0)
                            teks = Mid(isi, Len(meta) + 3)
                        End If
                        nAsli = AmbilAngka(meta, "baris_asli=")
                        nDas = AmbilAngka(itm.Body, "eligible sebanyak")
                        If nDas >= 0 And nAsli >= 0 And nDas <> nAsli Then
                            selisih = selisih & "- " & prog & " " & tgl & ": email DAS " & nDas & " transaksi, file " & nAsli & " baris" & vbCrLf
                        Else
                            Set rep = itm.ReplyAll
                            rep.Attachments.Add fixPath
                            rep.Attachments.Add baPath
                            If teks = "" Then teks = "Assalamualaikum Wr. Wb." & vbCrLf & vbCrLf & "(Isi email belum tersedia, mohon dilengkapi.)"
                            SisipkanTeks rep, teks
                            If uji Then
                                Do While rep.Recipients.Count > 0
                                    rep.Recipients.Remove 1
                                Loop
                                rep.Recipients.Add alamatUji
                                rep.Recipients.ResolveAll
                                rep.Subject = "[UJI COBA] " & rep.Subject
                                rep.Send
                            Else
                                If KIRIM_LANGSUNG Then rep.Send Else rep.Save
                                If itm.Categories = "" Then itm.Categories = KATEGORI Else itm.Categories = itm.Categories & ", " & KATEGORI
                                itm.Save
                            End If
                            dibuat = dibuat + 1
                        End If
                    End If
                End If
            End If
        End If
    Next itm

    Dim pesan As String
    If uji Then
        pesan = "UJI COBA: " & dibuat & " email terkirim ke " & alamatUji & vbCrLf & "Cek Inbox lo. Email DAS asli tidak ditandai." & vbCrLf
    Else
        pesan = IIf(KIRIM_LANGSUNG, "Balasan terkirim: ", "Draft balasan dibuat (cek di folder Drafts): ") & dibuat & vbCrLf
    End If
    pesan = pesan & _
            "Sudah pernah dibalas (dilewati): " & sudah & vbCrLf & _
            "Email 0 transaksi (tidak perlu dibalas): " & nol & vbCrLf
    If belumFix <> "" Then pesan = pesan & vbCrLf & "Belum ada file FIX untuk:" & vbCrLf & belumFix
    If selisih <> "" Then pesan = pesan & vbCrLf & "TIDAK dibuat karena jumlah tidak cocok (cek manual):" & vbCrLf & selisih
    MsgBox pesan, vbInformation, "KILAU"
End Sub

' ---------------------------------------------------------------------
Private Function AmbilTanggal(ByVal s As String) As String
    Dim p As Long, t As String
    p = InStr(1, s, "Periode Data", vbTextCompare)
    If p = 0 Then Exit Function
    t = Trim(Mid(s, p + Len("Periode Data")))
    t = Left(t, 10)
    If t Like "####-##-##" Then AmbilTanggal = t
End Function

Private Function CariFix(ByVal folder As String, ByVal prog As String, ByVal tgl As String) As String
    Dim f As String
    f = Dir(folder & "\*(FIX).xlsx")
    Do While f <> ""
        If InStr(1, f, prog, vbTextCompare) > 0 And InStr(1, f, tgl, vbTextCompare) > 0 And InStr(1, f, "multiposting", vbTextCompare) > 0 Then
            CariFix = folder & "\" & f
            Exit Function
        End If
        f = Dir()
    Loop
End Function

' Mengambil angka pertama setelah kata kunci; -1 jika tidak ada
Private Function AmbilAngka(ByVal s As String, ByVal kunci As String) As Long
    Dim p As Long, i As Long, c As String, d As String
    AmbilAngka = -1
    p = InStr(1, s, kunci, vbTextCompare)
    If p = 0 Then Exit Function
    For i = p + Len(kunci) To Len(s)
        c = Mid(s, i, 1)
        If c Like "#" Then
            d = d & c
        ElseIf d <> "" Then
            Exit For
        End If
        If i > p + Len(kunci) + 20 And d = "" Then Exit For
    Next i
    If d <> "" Then AmbilAngka = CLng(d)
End Function

Private Function BacaUtf8(ByVal jalur As String) As String
    Dim st As Object
    If Dir(jalur) = "" Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile jalur
    BacaUtf8 = st.ReadText
    st.Close
    If Left(BacaUtf8, 1) = ChrW(&HFEFF) Then BacaUtf8 = Mid(BacaUtf8, 2)
End Function

Private Sub SisipkanTeks(rep As Outlook.MailItem, ByVal teks As String)
    Dim h As String, p As Long, q As Long, blok As String
    teks = Replace(teks, "&", "&amp;")
    teks = Replace(teks, "<", "&lt;")
    teks = Replace(teks, ">", "&gt;")
    teks = Replace(teks, vbCrLf, "<br>")
    teks = Replace(teks, vbLf, "<br>")
    blok = "<div style=""font-family:Calibri,Arial,sans-serif;font-size:11pt"">" & teks & "</div><br>"
    h = rep.HTMLBody
    p = InStr(1, h, "<body", vbTextCompare)
    If p > 0 Then q = InStr(p, h, ">")
    If p > 0 And q > 0 Then
        rep.HTMLBody = Left(h, q) & blok & Mid(h, q + 1)
    Else
        rep.HTMLBody = blok & h
    End If
End Sub
