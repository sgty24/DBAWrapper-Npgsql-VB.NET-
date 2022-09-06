# DBAWrapper-Npgsql-VB.NET-
VB.NETで書かれたNpgsqlのラッパーです。
sample.vb の用にパラメーター指定を簡素に書けるようにしています。

# 制限
現在以下のパラメーターに対応しています。

・Int32
・Int32配列
・String
・String配列
・DateTimeEx(※)
・Boolean
・Boolean配列
・Decimal
・Decimal配列

※VB.NET標準の日付型はDateとDateTimeが区別されずDBパラメーターとして使用する際に型判別ができないため、DateTimeをラップしたクラスを作成しました。

# 動作環境
以下環境で動作確認しています。
- .NET Framework 4.8
- Npgsql 6.0.6
- PostgreSQL 14.5
