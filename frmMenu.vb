Public Class frmMenu
    Const MDBCONNECT As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Const WRHAND_PRG As String = "WrhsMain.exe"

    Const DIR_CONFIG As String = "henkan\"
    '仕訳伝票
    Const DIR_HENKAN As String = "henkan\"
    Const DIR_OK As String = "Ok\"
    Const DIR_INCSV As String = "分割\"
    Const DIR_BREAK As String = "中断\"
    '入金伝票
    Const DIR_INCSV_N As String = "分割_n\"
    Const DIR_BREAK_N As String = "中断_n\"
    '出金伝票
    Const DIR_INCSV_S As String = "分割_s\"
    Const DIR_BREAK_S As String = "中断_s\"

    Const CONFIGFILE As String = "Kanjo2kconfig.mdb"         '設定データベース
    Const PRGSIWAKE As String = "mntsiwake.exe"     '仕訳伝票プログラム
    Const PRGNYUKIN As String = "MntNyukin.exe"     '入金伝票プログラム
    Const PRGSHUKIN As String = "MntShukin.exe"     '出金伝票プログラム
    Const PRGKOTEI As String = "固定項目保守.exe"      '入金出金固定項目設定プログラム
    Const JOB_SIWAKE As String = "勘定奉行仕訳伝票"
    Const JOB_NYUKIN As String = "勘定奉行入金伝票"
    Const JOB_SHUKIN As String = "勘定奉行出金伝票"

    Dim JOBNAME_SIWAKE As String    '2014.09.02

    '入力データ有無フラグ
    '仕訳伝票
    Dim pblFlgINFILE As Boolean      '変換データ
    Dim pblFlgDIVFILE As Boolean      '分割後データ
    Dim pblFlgRecFILE As Boolean      '中断データ
    '入金伝票
    Dim pblFlgINFILE_N As Boolean      '変換データ
    Dim pblFlgDIVFILE_N As Boolean      '分割後データ
    Dim pblFlgRecFILE_N As Boolean      '中断データ
    '出金伝票
    Dim pblFlgINFILE_S As Boolean      '変換データ
    Dim pblFlgDIVFILE_S As Boolean      '分割後データ
    Dim pblFlgRecFILE_S As Boolean      '中断データ

    Dim pblInstPath As String
    Dim pblWinReaderPath As String  'WrHandインストール先

    '各EXE起動
    '引　数：EXEパス，エラーメッセージ，
    Private Function fncOpenExe(ByVal psExePath As String, ByVal psOpenErr As String) As Integer
        'Shell戻り値
        Dim vRetPrg As Object
        Dim ErrMsg As String
        Dim MSG_TITLE As String

        On Error GoTo ErrLine

        vRetPrg = 0
        fncOpenExe = 1
        vRetPrg = Interaction.Shell(psExePath, vbNormalFocus, True)

        'Shell起動時に終了まで待つを設定したためコメント()
        'Shell戻りがエラー発生()
        'If (vRetPrg = 0) Then
        '    ErrMsg = psOpenErr & "処理中に問題が発生しました。"
        '    MsgBox(ErrMsg, vbOKOnly + vbExclamation)
        '    Exit Function
        'End If

        'Shell起動プログラムの終了を待つ'''123
        'Call subShellEndWait(vRetPrg)

        Exit Function

        '
ErrLine:  'エラー
        MsgBox("[ " & psOpenErr & " ] " & "の処理に失敗しました。" & vbNewLine & ErrMsg, vbCritical)
        fncOpenExe = 0
    End Function

    Public Sub ErrMessage(ByVal Msg As String)
        '---エラー時の処理
        MsgBox(Msg & "にエラーが発生したため、処理を終了します。")
        End
    End Sub

    Private Sub FileSerch()
        '----------------------------------------------------------------
        '   ファイル有無チェック
        '----------------------------------------------------------------
        Dim myPath As String

        On Error GoTo ErrPrc

        '設定データベース
        If (Dir(pblInstPath & DIR_HENKAN & CONFIGFILE) = "") Then
            MsgBox("設定データベースがありません。" & vbNewLine & "ソフトを再インストールしてください。", vbOKOnly Or vbExclamation, "ＮＧ")
            End
        End If

        '---仕訳伝票の存在をチェック
        '中断ファイル
        pblFlgRecFILE = False
        myPath = Dir(pblInstPath & DIR_BREAK, vbDirectory)
        Do While myPath <> ""
            ' 現在のフォルダと親フォルダは無視します。
            If myPath <> "." And myPath <> ".." Then
                pblFlgRecFILE = True
                Exit Do
            End If
            myPath = Dir()
        Loop
        '分割ファイル
        pblFlgDIVFILE = True
        If (Dir(pblInstPath & DIR_INCSV & "*.csv") = "") Then
            pblFlgDIVFILE = False
        End If

        '---仕訳伝票ボタン状態
        cmdSScan.Enabled = True
        cmdSTyu.Enabled = True

        If (Dir(pblInstPath & DIR_HENKAN & PRGSIWAKE) = "") Then
            cmdSScan.Enabled = False
            cmdSTyu.Enabled = False
        Else
            If pblFlgRecFILE = False Then
                cmdSTyu.Enabled = False
            End If

        End If

        '---入金伝票の存在をチェック
        '中断ファイル
        pblFlgRecFILE_N = False
        myPath = Dir(pblInstPath & DIR_BREAK_N, vbDirectory)
        Do While myPath <> ""
            ' 現在のフォルダと親フォルダは無視します。
            If myPath <> "." And myPath <> ".." Then
                pblFlgRecFILE_N = True
                Exit Do
            End If
            myPath = Dir()
        Loop
        '分割ファイル
        pblFlgDIVFILE_N = True
        If (Dir(pblInstPath & DIR_INCSV_N & "*.csv") = "") Then
            pblFlgDIVFILE_N = False
        End If

        '---入金伝票ボタン状態
        cmdNScan.Enabled = True
        cmdNTyu.Enabled = True
        If (Dir(pblInstPath & DIR_HENKAN & PRGNYUKIN) = "") Then
            cmdNScan.Enabled = False
            cmdNTyu.Enabled = False
        Else
            If pblFlgRecFILE_N = False Then
                cmdNTyu.Enabled = False
            End If
        End If

        '---出金伝票の存在をチェック
        '中断ファイル
        pblFlgRecFILE_S = False
        myPath = Dir(pblInstPath & DIR_BREAK_S, vbDirectory)
        Do While myPath <> ""
            ' 現在のフォルダと親フォルダは無視します。
            If myPath <> "." And myPath <> ".." Then
                pblFlgRecFILE_S = True
                Exit Do
            End If
            myPath = Dir()
        Loop
        '分割ファイル
        pblFlgDIVFILE_S = True
        If (Dir(pblInstPath & DIR_INCSV_S & "*.csv") = "") Then
            pblFlgDIVFILE_S = False
        End If

        '---出金伝票ボタン状態
        cmdOScan.Enabled = True
        cmdOTyu.Enabled = True
        If (Dir(pblInstPath & DIR_HENKAN & PRGSHUKIN) = "") Then
            cmdOScan.Enabled = False
            cmdOTyu.Enabled = False
        Else
            If pblFlgRecFILE_S = False Then
                cmdOTyu.Enabled = False
            End If

        End If

        '---入出金固定項目ボタン状態
        If (Dir(pblInstPath & DIR_HENKAN & PRGKOTEI) = "") Then
            cmdKotei.Enabled = False
        Else
            cmdKotei.Enabled = True
        End If

        Exit Sub

ErrPrc:
        Call ErrMessage("初期ファイル検索中")

    End Sub

    Public Function GetPath()
        '----------------------------------------------------------------
        '   パス情報取得
        '----------------------------------------------------------------
        Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\FKDL", False)
        GetPath = ""
        If regkey Is Nothing Then
            Exit Function
        End If

        '展開して取得する
        GetPath = CStr(regkey.GetValue("InstDir"))

    End Function

    Public Function WrHandInstPath()
        '----------------------------------------------------------------
        '   インストールパス情報取得
        '----------------------------------------------------------------
        'Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\MediaDrive\帳票定義\INSTALL", False)
        Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\MediaDrive\WinReader HAND S", False)
        WrHandInstPath = ""
        If regkey Is Nothing Then
            Exit Function
        End If

        '展開して取得する
        'WrHandInstPath = CStr(regkey.GetValue("DIR"))
        WrHandInstPath = CStr(regkey.GetValue("path"))

    End Function

    Private Sub CmdDetail(ByVal mode As Integer, ByVal JobName As String, ByVal StartEXE As String)
        'ボタン処理共通記述 mode: 0:SCAN 1:修正のみ
        Dim cn As New ADODB.Connection             'データベース情報
        Dim RecSet As New ADODB.Recordset              'レコードセット
        Dim sSql As String
        Dim retVal As String

        If (mode = 0) Or (mode = 1) Then
            '---config.mdb rewrite
            'データベース接続
            cn.ConnectionString = MDBCONNECT & pblInstPath & DIR_CONFIG & CONFIGFILE   'データベース接続情報
            cn.Open()

            sSql = "select * from config"
            RecSet.Open(sSql, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            RecSet.Fields("sub2").Value = mode

            RecSet.Update()
            RecSet.Close()

            cn.Close()
            cn = Nothing
        End If

        Timer1.Enabled = False
        Me.Visible = False

        ' WinReaderHandS起動　2014/09/02
        If mode = 0 Then
            ' マウス ポインタを砂時計に変更します。
            Me.Cursor = Cursors.WaitCursor

            'retVal = fncOpenExe(pblWinReaderPath & "\\" & WRHAND_PRG & " """ & JobName & """ /h2", "伝票スキャン")

            Dim p As New System.Diagnostics.ProcessStartInfo()

            ' 起動するアプリケーションを設定する
            p.FileName = pblWinReaderPath & "\\" & WRHAND_PRG

            ' コマンドライン引数を設定する（WinReaderのJOB起動パラメーター）
            p.Arguments = " """ & JOBNAME_SIWAKE & """ /H2"

            ' WinReaderを起動します
            Dim hProcess As System.Diagnostics.Process
            hProcess = System.Diagnostics.Process.Start(p)

            ' WinReaderが終了するまで待機する
            hProcess.WaitForExit()

            ' マウス ポインタを元に戻します。
            Me.Cursor = Cursors.Default
        End If

        '---修正画面起動
        If mode = 2 Then
            '固定項目設定画面
        Else
            retVal = fncOpenExe(pblInstPath & DIR_HENKAN & StartEXE, JobName)
        End If

        '---メニュー再表示
        Call FileSerch()

        Me.Visible = True
        Timer1.Enabled = True

    End Sub

    Private Sub frmMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, "確認") = vbNo Then
            e.Cancel = True
            'Else
            '    End
        End If

    End Sub

    Private Sub frmMenu_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave

    End Sub


    Private Sub frmMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SFileName As String

        On Error GoTo ErrPrc

        '---------------------------------------------
        ' インスタンス実行中チェック
        '---------------------------------------------
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            MsgBox("このアプリケーションはすでに実行されています。", vbOKOnly Or vbExclamation, "ＮＧ")
            End
        End If

        '---------------------------------------------
        ' 作業ディレクトリ情報取得
        '---------------------------------------------
        'アプリケーションインストール先パス 2014/09/02
        'pblInstPath = GetPath()
        pblInstPath = My.MySettings.Default.appInstPath

        'WinReaderHandsインストール先パス 2014/09/02
        pblWinReaderPath = My.MySettings.Default.wrhInstPath

        'pblWinReaderPath = WrHandInstPath()
        'If pblWinReaderPath = "" Then
        '    'WinReaderのpathが取得できない場合
        '    MsgBox("WinReaderHandのインストール先を取得できませんでした。", vbOKOnly Or vbExclamation, "ＮＧ")
        '    End
        'End If

        '---------------------------------------------
        ' フォルダ作成
        '---------------------------------------------
        On Error Resume Next

        '振替伝票
        MkDir(pblInstPath & DIR_OK)
        MkDir(pblInstPath & DIR_INCSV)         '分割したＣＳＶの格納フォルダ 2004/6/24
        MkDir(pblInstPath & DIR_BREAK)         '中断伝票フォルダ 2004/6/24

        '入金伝票
        MkDir(pblInstPath & DIR_INCSV_S)       '分割したＣＳＶの格納フォルダ 2004/9/8
        MkDir(pblInstPath & DIR_BREAK_S)       '中断伝票フォルダ 2004/9/8

        '出金伝票
        MkDir(pblInstPath & DIR_INCSV_N)       '分割したＣＳＶの格納フォルダ 2004/9/8
        MkDir(pblInstPath & DIR_BREAK_N)       '中断伝票フォルダ 2004/9/8

        On Error GoTo ErrPrc

        '---2007.04.11 job&template移動
        If Dir(pblInstPath & DIR_HENKAN & "勘定奉行*.*") <> "" Then
            '-----仕訳伝票
            SFileName = "勘定奉行仕訳伝票"
            On Error Resume Next
            'JOBファイル移動
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".JOB", pblWinReaderPath & "\\JOB\\" & SFileName & ".JOB")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".JOB")
            'TEMPLATE移動
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".MRB", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".MRB")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".MRB")
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".tpl", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".tpl")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".tpl")
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".tpt", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".tpt")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".tpt")

            '-----入金伝票
            SFileName = "勘定奉行入金伝票"
            'JOBファイル移動
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".JOB", pblWinReaderPath & "\\JOB\\" & SFileName & ".JOB")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".JOB")
            'TEMPLATE移動
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".MRB", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".MRB")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".MRB")
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".tpl", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".tpl")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".tpl")
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".tpt", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".tpt")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".tpt")

            '-----出金伝票
            SFileName = "勘定奉行出金伝票"
            'JOBファイル移動
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".JOB", pblWinReaderPath & "\\JOB\\" & SFileName & ".JOB")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".JOB")
            'TEMPLATE移動
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".MRB", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".MRB")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".MRB")
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".tpl", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".tpl")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".tpl")
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".tpt", pblWinReaderPath & "\\TEMPLATE\\" & SFileName & ".tpt")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".tpt")

            'くせ字
            SFileName = "勘定奉行出金伝票くせ字"
            'JOBファイル移動
            FileCopy(pblInstPath & DIR_HENKAN & SFileName & ".JOB", pblWinReaderPath & "\\JOB\\" & SFileName & ".JOB")
            Kill(pblInstPath & DIR_HENKAN & SFileName & ".JOB")


            'MDBファイル移動
            FileCopy(pblInstPath & DIR_HENKAN & "wrh3.mdb", pblWinReaderPath & "\\wrh3.mdb")
            Kill(pblInstPath & DIR_HENKAN & "wrh3.mdb")

            On Error GoTo ErrPrc
        End If

        '----------------------------------------------
        '   タイマー発生間隔
        '----------------------------------------------
        Timer1.Interval = 5000

        '----------------------------------------------
        ' ファイル有無チェック
        '----------------------------------------------
        Call FileSerch()

        'WinReaderHandS Job名
        JOBNAME_SIWAKE = My.MySettings.Default.wrhJobName

        On Error GoTo 0

        Exit Sub

        ' エラー処理
ErrPrc:
        Call ErrMessage("初期設定中")

    End Sub

    Private Sub cmdSScan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSScan.Click
        CmdDetail(0, JOBNAME_SIWAKE, PRGSIWAKE)
    End Sub

    Private Sub cmdSTyu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSTyu.Click
        CmdDetail(1, JOBNAME_SIWAKE, PRGSIWAKE)
    End Sub

    Private Sub cmdNScan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNScan.Click
        CmdDetail(0, JOB_NYUKIN, PRGNYUKIN)
    End Sub

    Private Sub cmdNTyu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNTyu.Click
        CmdDetail(1, JOB_NYUKIN, PRGNYUKIN)
    End Sub

    Private Sub cmdOScan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOScan.Click
        CmdDetail(0, JOB_SHUKIN, PRGSHUKIN)
    End Sub

    Private Sub cmdOTyu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOTyu.Click
        CmdDetail(1, JOB_SHUKIN, PRGSHUKIN)
    End Sub

    Private Sub cmdKotei_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKotei.Click
        CmdDetail(1, "入出金伝票固定項目設定", PRGKOTEI)
    End Sub
End Class
