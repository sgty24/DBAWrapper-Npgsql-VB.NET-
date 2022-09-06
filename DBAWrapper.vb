Imports Npgsql
Imports NpgsqlTypes

''' ////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>日付型拡張クラス</summary>
''' <remarks>
''' VB.NET標準の日付型はDateとDateTimeが区別されずDBパラメーターとして使用する際に
''' 型判別ができない。これを解消するためDateTimeをラップしたクラスを作成する。
''' </remarks>
Public Class DateTimeEx
    '*******************************************************************************************
    '** メンバ変数 *****************************************************************************
    ''' <summary>時間データ有無</summary>
    Private m_with_time As Boolean

    '*******************************************************************************************
    '** プロパティ *****************************************************************************
    ''' ----------------------------------------------------------------------------------------
    ''' <summary>日付・時間データ</summary>
    Public Property Value As DateTime

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>時間データ有無プロパティ</summary>
    Public ReadOnly Property WithTime As Boolean
        Get
            Return m_with_time
        End Get
    End Property

    '*******************************************************************************************
    '** メンバ関数 *****************************************************************************
    ''' ----------------------------------------------------------------------------------------
    ''' <summary>デフォルトコンストラクタ</summary>
    Public Sub New()
        Value = New DateTime
        m_with_time = True
    End Sub
    
    ''' ----------------------------------------------------------------------------------------
    ''' <summary>コンストラクタ(値指定)</summary>
    ''' <param name="_val">日付・時間データ</param>
    ''' <param name="_with_time">時間データ有無</param>
    Public Sub New(ByVal _val As DateTime, ByVal _with_time As Boolean)
        Value = _val
        m_with_time = _with_time
    End Sub
End Class

''' ////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>DB接続管理クラス</summary>
Public Class DBAWrapper
    '*******************************************************************************************
    '** 継承 ***********************************************************************************
    Implements IDisposable

    '*******************************************************************************************
    '** メンバ変数 *****************************************************************************
    ''' <summary>PostgreSQL接続オブジェクト</summary>
    Private m_cn As New NpgsqlConnection

    ''' <summary>PostgreSQLコマンドオブジェクト</summary>
    Private m_cmd As NpgsqlCommand = Nothing

    ''' <summary>PostgreSQLレコードオブジェクト</summary>
    Private m_dr As NpgsqlDataReader = Nothing

    ''' <summary>パラメーターリスト</summary>
    Private m_param_list As New List(Of Npgsql.NpgsqlParameter)

    ''' <summary>パラメーターリスト</summary>
    Private m_is_begin_tran As Boolean = False
    Private disposedValue As Boolean

    '*******************************************************************************************
    '** メンバ関数 *****************************************************************************
    ''' ----------------------------------------------------------------------------------------
    ''' <summary>コンストラクタ</summary>
    Public Sub New()
        Me.m_cn.ConnectionString = My.Settings.connect_str
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>DB接続オープン</summary>
    Public Sub OpenConnect()
        Me.m_cn.Open()
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター付SQL実行(結果セットなし)</summary>
    ''' <param name="_sql">実行SQL</param>
    Public Sub Exec(ByVal _sql As String)
        CloseDataReader()
        Me.m_cmd = New NpgsqlCommand
        Me.m_cmd.Connection = Me.m_cn
        Me.m_cmd.CommandText = _sql
        For Each param As Npgsql.NpgsqlParameter In Me.m_param_list
            Me.m_cmd.Parameters.Add(param)
        Next
        Me.m_cmd.ExecuteNonQuery()
        Me.m_param_list.Clear()
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター付SQL実行(結果セットあり)</summary>
    ''' <param name="_sql">実行SQL</param>
    Public Function ExecSelect(ByVal _sql As String) As Npgsql.NpgsqlDataReader
        CloseDataReader()
        Me.m_cmd = New NpgsqlCommand
        Me.m_cmd.Connection = Me.m_cn
        Me.m_cmd.CommandText = _sql
        For Each param As Npgsql.NpgsqlParameter In Me.m_param_list
            Me.m_cmd.Parameters.Add(param)
        Next
        Me.m_dr = Me.m_cmd.ExecuteReader()
        Me.m_param_list.Clear()
        Return Me.m_dr
    End Function

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>データセットを取得する(パラメータあり)</summary>
    ''' <param name="_sql">>実行SQL</param>
    ''' <returns>データセット</returns>
    Public Function Fill(ByVal _sql As String) As System.Data.DataSet
        CloseDataReader()

        Dim da As New NpgsqlDataAdapter
        Dim ds As New System.Data.DataSet
        Me.m_cmd = New NpgsqlCommand(_sql, m_cn)

        For Each param As Npgsql.NpgsqlParameter In Me.m_param_list
            Me.m_cmd.Parameters.Add(param)
        Next

        da.SelectCommand = Me.m_cmd
        da.Fill(ds)

        da = Nothing
        Me.m_param_list.Clear()

        Return ds
    End Function

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>Npgsqlパラメーター生成</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    ''' <returns>Npgsqlパラメーター</returns>
    Private Function GetNpgsqlParam(ByVal _param As String, ByVal _value As Object, ByVal _set_null As Boolean) As NpgsqlParameter
        Dim db_val As NpgsqlParameter
        Dim tmp_type As Type = _value.GetType()
        Dim array_mask As NpgsqlDbType = 0
        If tmp_type.IsArray() Then
            array_mask = NpgsqlDbType.Array
        End If

        'パラメータ型設定
        If (tmp_type = GetType(Int32)) Or (tmp_type = GetType(Int32())) Then
            db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Integer)
            db_val.Value = _value
        ElseIf (tmp_type = GetType(String)) Or (tmp_type = GetType(String())) Then
            db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Varchar)
            db_val.Value = _value
        ElseIf tmp_type = GetType(DateTimeEx) Then
            Dim tmp_dtex As DateTimeEx = DirectCast(_value, DateTimeEx)
            If tmp_dtex.WithTime Then
                db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Timestamp)
            Else
                db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Date)
            End If
            db_val.Value = tmp_dtex.Value
        ElseIf tmp_type = GetType(DateTimeEx()) Then
            Dim tmp_dtex As DateTimeEx() = DirectCast(_value, DateTimeEx())
            Dim tmp_array_org As DateTimeEx() = DirectCast(_value, DateTimeEx())
            Dim tmp_array As Date()
            If tmp_dtex(0).WithTime Then
                db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Timestamp)
            Else
                db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Date)
            End If
            ReDim tmp_array(0 To tmp_array_org.Length - 1)
            For i = 0 To tmp_array_org.Length - 1
                tmp_array(i) = tmp_array_org(i).Value
            Next
            db_val.Value = tmp_array
        ElseIf (tmp_type = GetType(Boolean)) Or (tmp_type = GetType(Boolean())) Then
            db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Boolean)
            db_val.Value = _value
        ElseIf (tmp_type = GetType(Decimal)) Or (tmp_type = GetType(Decimal())) Then
            db_val = New NpgsqlParameter(_param, array_mask Or NpgsqlDbType.Numeric)
            db_val.Value = _value
        Else
            Throw New Exception("ParamAddNext：未定義のデータ型です。")
        End If

        'NULL指定なら値上書き
        If _set_null Then
            db_val.Value = System.DBNull.Value
        End If

        Return db_val
    End Function

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/Int32)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As Int32, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/String)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As String, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/DateTimeEx)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As DateTimeEx, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/Boolean)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As Boolean, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/Decimal)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As Decimal, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/Int32)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As Int32, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/String)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As String, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/DateTimeEx)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As DateTimeEx, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/Boolean)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As Boolean, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/Decimal)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As Decimal, Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/Int32配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As Int32(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/String配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As String(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/DateTimeEx配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As DateTimeEx(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/Boolean配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As Boolean(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回/Decimal配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddFirst(ByVal _param As String, ByVal _value As Decimal(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Clear()
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/Int32配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As Int32(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/String配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As String(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/DateTimeEx配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As DateTimeEx(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/Boolean配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As Boolean(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>パラメーター追加(初回以降/Decimal配列)</summary>
    ''' <param name="_param">パラメーター名</param>
    ''' <param name="_value">値</param>
    ''' <param name="_set_null">NULL値をセットするか？</param>
    Public Sub ParamAddNext(ByVal _param As String, ByVal _value As Decimal(), Optional ByVal _set_null As Boolean = False)
        Me.m_param_list.Add(Me.GetNpgsqlParam(_param, _value, _set_null))
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>結果セットを閉じる</summary>
    Private Sub CloseDataReader()
        Try
            If Not Me.m_dr Is Nothing Then
                Me.m_dr.Close()
                Me.m_cmd = Nothing
                Me.m_dr = Nothing
            End If
        Catch _ex As Exception
            '無視
        End Try
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>トランザクション開始</summary>
    Public Sub BeginTran()
        Me.Exec("begin transaction;")
        m_is_begin_tran = True
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>トランザクションコミット</summary>
    Public Sub CommitTran()
        Me.Exec("commit;")
        m_is_begin_tran = False
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>トランザクションロールバック</summary>
    Public Sub RollbackTran()
        Try
            If Me.m_cn.State = ConnectionState.Open Then
                m_is_begin_tran = False
                Me.Exec("rollback;")
            End If
        Catch _ex As Exception
            '無視
        End Try
    End Sub

    ''' ----------------------------------------------------------------------------------------
    ''' <summary>オブジェクト廃棄処理</summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        If m_is_begin_tran Then
            Me.RollbackTran()
        End If
        CloseDataReader()
        Me.m_cmd = Nothing
        If Me.m_cn.State = ConnectionState.Open Then
            Me.m_cn.Close()
        End If

        GC.SuppressFinalize(Me)
    End Sub
End Class
