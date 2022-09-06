Module Main

    Sub Main()

        Try
            Using dba As New DBAWrapper
                Dim dr_ref As Npgsql.NpgsqlDataReader

                'DBオープン
                dba.OpenConnect()

                dba.ParamAddFirst("@v_date", New DateTimeEx(DateTime.Now, True))
                dba.ParamAddNext("@v_name", "t%")
                dr_ref = dba.ExecSelect("SELECT " &
                                          "sample_id," &
                                          "sample_date, " &
                                          "sample_name " &
                                        "FROM " &
                                          "sample_table " &
                                        "WHERE " &
                                          "sample_date<=@v_date AND " &
                                          "sample_name LIKE @v_name;")

                '全件コンソールへ表示
                Dim rec_num As Int64 = 0
                While dr_ref.Read()
                    rec_num += 1
                    Console.WriteLine(DirectCast(dr_ref("sample_id"), String) & ", " &
                                      DirectCast(dr_ref("sample_date"), DateTime) & ", " &
                                      DirectCast(dr_ref("sample_name"), String))
                End While
                Console.WriteLine("該当件数:" & rec_num.ToString())
            End Using
        Catch _ex As Exception
            'エラーハンドリング
            Console.WriteLine(_ex.Message)
        End Try
        Console.ReadKey()
    End Sub

End Module
