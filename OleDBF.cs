using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.Data.OleDb;
using System.IO;

namespace CardPerso
{
    
     //Класс работы с данными
     public class DataBaseHelper
     {
          //Конструктор
          public DataBaseHelper()
           {
           }

          //Строка подключения к базе дынных
          private static string GetConnectionString()
           {
               //return string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", FileHelper.GetDirectoryData());

               return string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=dBASE IV;Data Source={0}", FileHelper.GetDirectoryData());

           }

          //OleDbConnection
          private static OleDbConnection GetOleDbConnection()
           {
                OleDbConnection m_conn = new OleDbConnection();
                m_conn.ConnectionString = DataBaseHelper.GetConnectionString();

               return m_conn;
           }

          //Заполнение DataTable
          public static DataTable GetDataTable(string p_query, OleDbParameter[] p_pr)
          {
              DataTable m_dt = new DataTable();
              OleDbConnection m_conn = DataBaseHelper.GetOleDbConnection();
              OleDbCommand m_cmd = new OleDbCommand();
              OleDbDataAdapter m_da = new OleDbDataAdapter();

              try
              {
                  m_conn.Open();
              }
              catch (OleDbException e)
              {
                  throw new Exception(e.Message);
              }
              catch (Exception e1)
              {
                  throw new Exception(e1.Message);
              }

              m_cmd.Connection = m_conn;
              m_cmd.CommandText = p_query;

              if (p_pr != null)
              {
                  for (int i = 0; i < p_pr.Length; i++)
                  {
                      m_cmd.Parameters.Add(p_pr[i]);
                  }
              }

              m_da.SelectCommand = m_cmd;

              try
              {
                  m_da.Fill(m_dt);
              }
              catch (OleDbException e)
              {
                  throw new Exception(e.Message);
              }
              catch (Exception e1)
              {
                  throw new Exception(e1.Message);
              }
              finally
              {
                  m_cmd.Cancel();
                  m_conn.Close();
              }

              return m_dt;
          }
              
          //Выполнение произвольного запроса
          public static bool ExecuteQuery(string p_query, OleDbParameter[] p_pr)
           {
                bool m_result = true;
                OleDbConnection m_conn = DataBaseHelper.GetOleDbConnection();
                OleDbCommand m_cmd = new OleDbCommand();

               try 
                {
                     m_conn.Open();
                }
                catch (OleDbException e)
                {
                     throw new Exception(e.Message);
                }
               catch (Exception e1)
               {
                   throw new Exception(e1.Message);
               }

                m_cmd.Connection = m_conn;
                m_cmd.CommandText = p_query;

               if (p_pr != null)
                {
                     for (int i = 0;i<p_pr.Length;i++)
                     {
                          m_cmd.Parameters.Add(p_pr[i]);
                     }
                }

               try
                {
                     m_cmd.ExecuteNonQuery();
                }
                catch (OleDbException e)
                {
                     m_result = false;
                     throw new Exception(e.Message);
                }
               catch (Exception e1)
               {
                   throw new Exception(e1.Message);
               }
                finally
                {
                     m_cmd.Cancel();
                     m_conn.Close();
                     m_conn.Dispose();
                
                }

               return m_result;
           }
     }
   
      
      
      //Класс работы с файлами и папками
     public class FileHelper
      {
          private static string g_datafolder = "";
          private static string g_appfolder = Directory.GetCurrentDirectory();

          public static void initDefault()
          {
            g_datafolder = "";
            g_appfolder = Directory.GetCurrentDirectory();
          }

          public static void SetDataFolder(String dataFolder,bool copy)
          {
              FolderExist(null, dataFolder, true);
              if (copy)
              {
                  string ds=GetDirectoryData();
                  g_datafolder = dataFolder;
                  string dt=GetDirectoryData();
                  string[] files = Directory.GetFiles(ds, "*.dbf");
                  foreach (string f in files)
                  {
                      string t=f.Replace(ds,dt);
                      if(f!=t) File.Copy(f, t, true);
                  }
              }
              else g_datafolder = dataFolder;
          }

          //Конструктор 
          public FileHelper()
          {
          
          }

          //Проверка существования каталога 
          public static bool FolderExist(string p_path, string p_folder, bool p_create)
          {
               if (p_path == null)
                     p_path = GetDirectoryApp();

               if (p_folder == null)
                     p_folder = FileHelper.g_datafolder;

               string m_fullpath = string.Format(@"{0}\{1}", p_path, p_folder);
               bool m_result = Directory.Exists(m_fullpath);

               if (!m_result && p_create)
                     Directory.CreateDirectory(m_fullpath);

               return m_result;
          }

         //Директория приложения 
          private static string GetDirectoryApp()
          {
                return g_appfolder;
          }

          public static bool SetDirectoryApp(string dirApp)
          {
              if (Directory.Exists(dirApp))
              {
                  g_appfolder = dirApp;
                  if (g_appfolder[g_appfolder.Length - 1] == '\\')
                      g_appfolder = g_appfolder.Remove(g_appfolder.Length - 1);
                  return true;
              }
              return false;
          }

         //Директория базы данных 
          public static string GetDirectoryData()
          {
              if(g_datafolder.Length>0)  return string.Format(@"{0}\{1}\", FileHelper.GetDirectoryApp(), FileHelper.g_datafolder);
              else return string.Format(@"{0}\", FileHelper.GetDirectoryApp());
          }

         //Проверка существования файла 
          public static bool FileExists(string p_file)
          {
                return File.Exists(p_file);
          }

          private static void setCode(string filename, byte wrbyte)
          {
              BinaryWriter bwrite = new BinaryWriter(File.Open(GetDirectoryData() + filename, FileMode.Open));
              bwrite.Seek(29, SeekOrigin.Begin);
              bwrite.Write(wrbyte);
              bwrite.Close();
          }

          public static void rename(string fileName, string newfileName)
          {
              try
              {
                  File.Move(GetDirectoryData() + fileName, GetDirectoryData() + newfileName);

              }
              catch (Exception) { }
          }

          public static void setCode866(string filename)
          {
              setCode(filename, 101);
          }
          public static void setCodeDefault(string filename)
          {
              setCode(filename, 0);
          }
          public static string getFullName(string filename)
          {
              return GetDirectoryData() + filename;
          }
          public static void DeleteFiles(String path, bool delDir)
          {
              string[] files = Directory.GetFiles(path);
              foreach (string f in files)
              {
                  try
                  {
                      File.Delete(f);
                  }
                  catch (Exception)
                  {

                  }
              }
              if (delDir)
              {
                  try
                  {
                      Directory.Delete(path, false);
                  }
                  catch (Exception)
                  {

                  }
              }
              
          }
      }

}
