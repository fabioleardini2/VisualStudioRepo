using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace TestXML
{
    public class TesteObjeto
    {
        public string nome;
        public int quantidade;
        public DateTime datacriacao;
    }
    
    public class Program
    {
        public static void Main(string[] args)
        {
            //var x = new Program();
            //x.XlsToXml();
            
            //var r = new Program();
            //r.ReadXml();

            var r = new Program();
            r.DataBaseToXml();
        }

        public void XlsToXml()
        {
            string txtConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\fabio.leardini\Desktop\Umail - Catálogo de Produtos - Importação - Copy.xlsx;Extended Properties=""Excel 8.0;HDR=No;IMEX=1"";";

            DataTable dt = new DataTable();
            OleDbConnection conn = new OleDbConnection(txtConnString);
            OleDbDataAdapter adapter = new OleDbDataAdapter();

            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();

                dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt.Rows.Count != 1)
                    throw new Exception("Quantidade de abas incorreta!");

                IEnumerable<DataRow> rows = (dt.Rows.Cast<DataRow>());
                string sheetName = rows.First()["TABLE_NAME"].ToString();
                dt.Reset();
                adapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", conn);
                adapter.Fill(dt);

                object[] header = dt.Rows[0].ItemArray;

                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    XmlElement element = xmlDoc.CreateElement("Product");
                    xmlDoc.AppendChild(element);

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        element.SetAttribute(header[j].ToString(), dt.Rows[i].ItemArray[j].ToString());
                    }

                    Console.WriteLine(Convert.ToString(xmlDoc.InnerXml));
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.ReadLine();

                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        public void DataBaseToXml()
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection(@"Password=qwert123@;Persist Security Info=True;User ID=umailngzuser;Initial Catalog=CampaignTemplateBuilder;Data Source=BRJODBS01");
            
            SqlCommand command = new SqlCommand();
            command.CommandType = CommandType.Text;
            command.CommandText = "SELECT TOP 1000 * FROM XMLProduct ORDER BY Id DESC";
            command.Connection = conn;

            SqlDataAdapter ad = new SqlDataAdapter(command);
            
            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();

                ad.Fill(dt);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    XmlElement element = xmlDoc.CreateElement("Product");
                    xmlDoc.AppendChild(element);

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        element.SetAttribute(dt.Columns[j].ColumnName, dt.Rows[i].ItemArray[j].ToString().Replace("'", ""));
                    }

                    conn.Close();
                    conn.Dispose();

                    Insert(Convert.ToInt32(dt.Rows[i].ItemArray[0]), xmlDoc.InnerXml);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.ReadLine();
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        public void Insert(int id, string value)
        {
            using (SqlConnection conn = new SqlConnection(@"Password=qwert123@;Persist Security Info=True;User ID=umailngzuser;Initial Catalog=UmailNG_AmericanasTeste;Data Source=BRJODBS01"))
            {
                using (SqlCommand command = new SqlCommand())
                {
                    string date = string.Concat(DateTime.Now.Year, '-', DateTime.Now.Month, '-', DateTime.Now.Day, " ", DateTime.Now.Hour, ':', DateTime.Now.Minute, ':', DateTime.Now.Second, ".", DateTime.Now.Millisecond);
                    SqlDataReader dr;

                    command.CommandType = CommandType.Text;
                    command.CommandText = string.Format("SELECT * FROM NGZProduct WHERE ProductId = {0}", id);
                    command.Connection = conn;

                    conn.Open();
                    dr = command.ExecuteReader();

                    if (!dr.HasRows)
                    {
                        dr.Close();
                        command.CommandType = CommandType.Text;
                        command.CommandText = string.Format("INSERT INTO NGZProduct VALUES ({0}, '{1}', '{2}')", id, value, date);
                        command.Connection = conn;
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public void ReadXml()
        {
            string connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=UmailNG_NGZDesenv;Data Source=BRJODBS02";

            OleDbConnection conn = new OleDbConnection(connectionString);
            OleDbDataAdapter connAdp = new OleDbDataAdapter("SELECT * FROM NGZProduct WHERE ProductId = 116120267", conn);

            DataTable dt = new DataTable();

            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();                

                connAdp.Fill(dt);

                string xml = dt.Rows[0].ItemArray[1].ToString();

                if (!xml.Contains(header))
                    xml = string.Concat(header, xml, footer);

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);

                XmlNamespaceManager nameSpace = new XmlNamespaceManager(xmlDoc.NameTable);
                nameSpace.AddNamespace("g", "http://base.google.com/ns/1.0");
                nameSpace.AddNamespace("c", "http://base.google.com/cns/1.0");

                var node = xmlDoc.SelectSingleNode("/rss/channel/item/title", nameSpace);

                Console.WriteLine(xmlDoc.InnerXml.ToString());
                Console.ReadLine();                
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.ReadLine();

                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        public string header = string.Concat("<?xml version=", @"""", "1.0", @"""", "?>", "<rss version=", @"""", "2.0", @"""", " xmlns:g=", @"""", "http://base.google.com/ns/1.0", @"""", "><channel><item>");
        public string footer = "</item></channel></rss>";
    }
}
