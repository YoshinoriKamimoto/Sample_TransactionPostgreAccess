using System.Linq;
using System.Xml.Linq;
using System.Data.SqlClient;
using Npgsql;
using System.Data.OleDb;

internal class Program
{
    /// <summary>
    /// テストメソッド
    /// </summary>
    private static void Main(string[] args)
    {
        Console.WriteLine("プログラム開始\nConnectPostgreDBメソッド呼び出し");
        ConnectPostgreDB();
        Console.WriteLine("プログラム終了");
    }

    // PostgreDBに接続するメソッド
    private static void ConnectPostgreDB()
    {
        // 接続文字列
        string connectionStr = "Server=localhost;Port=5432;Username=postgres;Password=test;Database=test;";

        // 実行するクエリ
        string queryStr = "UPDATE ";
        queryStr += "public.test ";
        queryStr += "SET ";
        queryStr += "id = '2';";

        try
        {
            // PostgreDB接続用のインスタンスを生成
            using (NpgsqlConnection connection = new NpgsqlConnection(connectionStr))
            {
                // PostgreDB接続開始
                connection.Open();
                Console.WriteLine("PostgreDB接続成功");


                // トランザクションの開始
                using (NpgsqlTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // クエリ実行用のインスタンスを生成
                        using (NpgsqlCommand command = new NpgsqlCommand(queryStr, connection))
                        {
                            // クエリ実行
                            Console.WriteLine("PostgreDBクエリ実行");
                            command.ExecuteNonQuery();
                            Console.WriteLine("PostgreDBクエリ実行成功");

                            // AccessDBへ接続してクエリを実行
                            Console.WriteLine("ConnectAccessDBメソッド呼び出し");
                            ConnectAccessDB();
                            Console.WriteLine("AccessDBクエリ実行成功");

                            // トランザクションのコミット
                            transaction.Commit();
                            Console.WriteLine("トランザクションのコミット成功");
                        }
                    }
                    // 例外処理
                    catch (Exception ex2)
                    {
                        Console.WriteLine("クエリ実行エラー\n" + ex2 + "\nロールバック実行");
                        try
                        {
                            // ロールバック
                            transaction.Rollback();
                            Console.WriteLine("ロールバック成功");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ロールバック失敗\n" + ex);
                        }
                    }
                }
            }
        }
        // 例外処理
        catch (Exception ex)
        {
            Console.WriteLine("PostgreDB接続エラー\n" + ex);
        }
    }


    // AccessDBに接続するメソッド
    private static void ConnectAccessDB()
    {
        // 接続文字列
        string connectionStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/Users/kamimoto/Desktop/work/tmp/Sample/AccessDB/testDB.accdb;";

        // 実行するクエリ
        string queryStr = "INSERT INTO ";
        queryStr += "test (id) ";
        queryStr += "VALUES ";
        queryStr += "('1');";

        try
            {
                // AccessDB接続用のインスタンスを生成
                using (OleDbConnection connection = new OleDbConnection(connectionStr))
                {
                    // AccessDB接続開始
                    connection.Open();
                    Console.WriteLine("AccessDB接続成功");

                    // クエリ実行用のインスタンスを生成
                    using (OleDbCommand command = new OleDbCommand(queryStr, connection))
                    {
                        // クエリを実行
                        Console.WriteLine("AccessDBクエリ実行");
                        command.ExecuteNonQuery();
                        Console.WriteLine("AccessDBクエリ実行成功");
                    }
                }
            }
            // 例外処理
            catch
            {
                // 呼び出し元に例外を投げる
                throw;
            }
    }
}