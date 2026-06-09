' ── SAP Auto Login ────────────────────────────────────────────────────────────
' Baca kredensial dari .env di folder project root, lalu login ke SAP.
' Dipanggil sebelum macro CAUFV/AUFM/ZPPCPFINT_GRAUTO.

' ── Baca .env ──────────────────────────────────────────────────────────────────
Dim scriptDir : scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
Dim projDir   : projDir   = Left(scriptDir, InStrRev(Left(scriptDir, Len(scriptDir) - 1), "\"))
Dim envPath   : envPath   = projDir & ".env"

Dim env : Set env = CreateObject("Scripting.Dictionary")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(envPath) Then
    MsgBox "File .env tidak ditemukan: " & envPath & Chr(13) & "Salin .env.example menjadi .env dan isi kredensial.", vbCritical, "SAP Login"
    WScript.Quit 1
End If

Dim ts : Set ts = fso.OpenTextFile(envPath, 1)
Dim line, parts
Do While Not ts.AtEndOfStream
    line = Trim(ts.ReadLine())
    If Left(line, 1) <> "#" And InStr(line, "=") > 0 Then
        parts = Split(line, "=", 2)
        env(Trim(parts(0))) = Trim(parts(1))
    End If
Loop
ts.Close

Dim sapClient : sapClient = env("SAP_CLIENT")
Dim sapUser   : sapUser   = env("SAP_USER")
Dim sapPass   : sapPass   = env("SAP_PASS")
Dim sapSystem : sapSystem = env("SAP_SYSTEM")

If sapUser = "" Or sapPass = "" Then
    MsgBox "SAP_USER atau SAP_PASS kosong di file .env.", vbCritical, "SAP Login"
    WScript.Quit 1
End If

' ── Hubungkan ke SAP GUI Scripting Engine ──────────────────────────────────────
Dim SapGuiAuto, application, connection, session

On Error Resume Next
Set SapGuiAuto  = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
On Error GoTo 0

If Not IsObject(application) Then
    MsgBox "SAP Logon tidak berjalan. Buka SAP Logon terlebih dahulu.", vbCritical, "SAP Login"
    WScript.Quit 1
End If

' ── Buka koneksi jika belum ada ────────────────────────────────────────────────
If application.Connections.Count = 0 Then
    Set connection = application.OpenConnection(sapSystem, True)
    If Not IsObject(connection) Then
        MsgBox "Gagal membuka koneksi ke sistem: " & sapSystem, vbCritical, "SAP Login"
        WScript.Quit 1
    End If
Else
    Set connection = application.Children(0)
End If

Set session = connection.Children(0)

' ── Tangani dialog auto-logout sebelum cek status login ───────────────────────
' "SAP GUI for Windows 740: P01: auto logout" adalah dialog Windows milik SAP
' GUI client, bukan ABAP popup — tidak bisa dikontrol via session.sendVKey.
' Gunakan WScript.Shell.AppActivate + SendKeys untuk klik No.
Dim closedDialog : closedDialog = False
Dim wshDlg : Set wshDlg = CreateObject("WScript.Shell")
Dim dlgAttempt : dlgAttempt = 0
Do While dlgAttempt < 5
    If wshDlg.AppActivate("SAP GUI for Windows 740") Then
        WScript.Sleep 400
        wshDlg.SendKeys "n"           ' coba tombol No (dialog Yes/No)
        WScript.Sleep 400
        ' Jika dialog masih ada setelah "n", berarti dialog hanya punya OK
        If wshDlg.AppActivate("SAP GUI for Windows 740") Then
            WScript.Sleep 200
            wshDlg.SendKeys "{ENTER}" ' tekan OK untuk dialog OK-only
            WScript.Sleep 400
        End If
        closedDialog = True
        dlgAttempt = dlgAttempt + 1
    Else
        dlgAttempt = 5                ' tidak ada dialog lagi, keluar loop
    End If
Loop
Set wshDlg = Nothing

If closedDialog Then
    ' SAP perlu waktu untuk tampilkan login screen setelah session di-terminate.
    ' Re-acquire session agar tidak pakai object lama yang sudah stale.
    WScript.Sleep 2500
    On Error Resume Next
    Set session = connection.Children(0)
    Err.Clear : On Error GoTo 0
Else
    WScript.Sleep 500
End If

' ── Cek apakah sudah login (layar utama) atau masih di login screen ────────────
Dim onLoginScreen : onLoginScreen = False
On Error Resume Next
Dim testField : Set testField = session.findById("wnd[0]/usr/txtRSYST-MANDT")
If Err.Number = 0 Then onLoginScreen = True
Err.Clear
On Error GoTo 0

If Not onLoginScreen Then
    ' Sudah login, tidak ada popup — langsung lanjut
    WScript.Quit 0
End If

' ── Isi form login ─────────────────────────────────────────────────────────────
If sapClient <> "" Then
    session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = sapClient
End If
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = sapUser
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = sapPass
session.findById("wnd[0]").sendVKey 0

WScript.Sleep 2000

' ── Tangani dialog "multiple logon" jika muncul ────────────────────────────────
On Error Resume Next
Dim multiLogon : Set multiLogon = session.findById("wnd[1]")
Dim multiLogonMuncul : multiLogonMuncul = (Err.Number = 0)
Err.Clear
On Error GoTo 0

If multiLogonMuncul Then
    ' Batalkan login — user SAP sedang dipakai di sesi lain
    session.findById("wnd[1]").sendVKey 12  ' F12 = Cancel
    WScript.Quit 2
End If
