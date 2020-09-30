using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace CardPerso
{
    
    public class Ozekimessageout
    {
	    public int id { get; set; }
	    public string sender { get; set; }
	    public string receiver { get; set; }
	    public string msg { get; set; }
	    public DateTime senttime { get; set; }
	    public DateTime receivedtime { get; set; }
	    public string operators { get; set; }
	    public string msgtype { get; set; }
	    public string reference { get; set; }
	    public string status { get; set; }
	    public string errormsg { get; set; }

        public Ozekimessageout()
        {
            senttime = DateTime.MinValue;
            receivedtime = DateTime.MinValue;
        }
        public Ozekimessageout(string receiver, string msg)
        {
            senttime = DateTime.MinValue;
            receivedtime = DateTime.MinValue;
            this.receiver = receiver;
            this.msg = msg;
            msgtype = "SMS:TEXT";
            status = "send";
        }
    }

    public class Ozekidata
    {
        public int id { get; set; }
        public int id_doc { get; set; }
        public int id_card { get; set; }
        public int id_branchcard { get; set; }
        public string last4 { get; set; }
        public string company { get; set; }
        public string bin { get; set; }
        public string prefix_ow { get; set; }
        public string fio { get; set; }
        public string adress { get; set; }
        public Ozekimessageout msgout { get; set; }
    }
    
    public class Ozeki
    {
        private SqlConnection conn = null;
        
        
        public Ozeki(SqlConnection conn)
        {
            this.conn = conn;
        }

        public Ozekidata[] readData(int id_doc)
        {
            // вызовов нет, но не знаю зачем Олег ее писал. Пока что не стираю
            Ozekidata[] rd = null;
            //SqlCommand cmdSelect = new SqlCommand(string.Format("select OzekiData.[id]," +
            //                                                    "OzekiData.[id_doc]," +
            //                                                    "OzekiData.[id_card]," +
            //                                                    "Cards.[Phone]," +
            //                                                    "Cards.[id_branchcard]," +
            //                                                    "OzekiData.[id_ozeki]," +
            //                                                    "OzekiMessageOut.[sender]," +
            //                                                    "OzekiMessageOut.[receiver]," +
            //                                                    "OzekiMessageOut.[msg]," + 
            //                                                    "OzekiMessageOut.[senttime]," + 
            //                                                    "OzekiMessageOut.[receivedtime]," + 
            //                                                    "OzekiMessageOut.[operator]," + 
            //                                                    "OzekiMessageOut.[msgtype]," + 
            //                                                    "OzekiMessageOut.[reference]," + 
            //                                                    "OzekiMessageOut.[status]," + 
            //                                                    "OzekiMessageOut.[errormsg]" + 
            //                                                    " from OzekiData" +
            //                                                    " join OzekiMessageOut on OzekiData.id_ozeki = OzekiMessageOut.id" +
            //                                                    " join Cards on Cards.id = OzekiData.id_card where OzekiData.[id_doc] = {0}", id_doc), conn);
            //using (SqlDataReader dr = cmdSelect.ExecuteReader())
            //{
            //    if (dr.HasRows)
            //    {
            //        List<Ozekidata> lod = new List<Ozekidata>();
            //        while (dr.Read() == true)
            //        {
            //            Ozekidata r = new Ozekidata();
            //            r.msgout = new Ozekimessageout();

            //            r.id = Convert.ToInt32(dr["id"]);
            //            r.id_doc = Convert.ToInt32(dr["id_doc"]);
            //            r.id_card = Convert.ToInt32(dr["id_card"]);
            //            r.id_branchcard = Convert.ToInt32(dr["id_branchcard"]);
            //            r.msgout.id = Convert.ToInt32(dr["id_ozeki"]);
                        

            //            if (dr["sender"] != DBNull.Value) r.msgout.sender = (string)dr["sender"];
            //            if (dr["receiver"] != DBNull.Value) r.msgout.receiver = (string)dr["receiver"];
            //            if (dr["msg"] != DBNull.Value) r.msgout.msg = (string)dr["msg"];
            //            if (dr["senttime"] != DBNull.Value) r.msgout.senttime = Convert.ToDateTime(dr["senttime"]);
            //            if (dr["receivedtime"] != DBNull.Value) r.msgout.receivedtime = Convert.ToDateTime(dr["receivedtime"]);
            //            if (dr["operator"] != DBNull.Value) r.msgout.operators = (string)dr["operator"];
            //            if (dr["msgtype"] != DBNull.Value) r.msgout.msgtype = (string)dr["msgtype"];
            //            if (dr["reference"] != DBNull.Value) r.msgout.reference = (string)dr["reference"];
            //            if (dr["status"] != DBNull.Value) r.msgout.status = (string)dr["status"];
            //            if (dr["errormsg"] != DBNull.Value) r.msgout.errormsg = (string)dr["errormsg"];
            //            lod.Add(r);
            //        }
            //        rd = lod.ToArray();

            //    }
            //    else throw new Exception("readData: документ с идентификатором " + id_doc.ToString() + " не найден");
            //}
            return rd;
        }

        public void deleteData(int id_doc)
        {
            //16.10.18 - убрали удаление из ozekimessageout для предотвращения ситуации: оператор снял подтверждение приема, когда sms еще не ушла
            return;
            SqlTransaction tx = null;
            try
            {
                tx = conn.BeginTransaction();
                //SqlCommand cmdDelete = new SqlCommand(string.Format("delete from ozekimessageout where id in (select id_ozeki from ozekidata where id_doc={0})\r\ndelete from ozekidata where id_doc={0}", id_doc), conn);
                SqlCommand cmdDelete = new SqlCommand(string.Format("delete from ozekimessageout where id in (select id_ozeki from ozekidata where id_doc={0})", id_doc), conn);
                cmdDelete.CommandTimeout = conn.ConnectionTimeout;
                cmdDelete.Transaction = tx;
                cmdDelete.ExecuteNonQuery();
                tx.Commit();
            }
            catch (Exception e)
            {
                tx.Rollback();
                throw new Exception("deleteData: " + e.Message);
            }
        }

        public Ozekidata[] newData(int id_doc)
        {
            Ozekidata[] rd = null;
            try
            {
               
                SqlCommand SelectCommand = new SqlCommand(string.Format("select Cards_StorageDocs.id_doc," + 
                                                                        "Cards_StorageDocs.id_card," +
                                                                        "Cards.pan," +
                                                                        "Cards.company," + 
                                                                        "Cards.phone," +
                                                                        "Cards.fio," +
                                                                        "Cards.id_branchcard," + 
                                                                        "Branchs.adress," +
                                                                        "Products_Banks.bin," + 
                                                                        "Products.prefix_ow" + 
                                                                        " from Cards_StorageDocs" +  
                                                                        " join Cards on Cards.id = Cards_StorageDocs.id_card" +
                                                                        " join Branchs on Cards.id_BranchCard = Branchs.id" +  
                                                                        " join Products_Banks on Cards.id_prb = Products_banks.id" +
                                                                        " join Products on Products.id = Products_Banks.id_prod" + 
                                                                        " where id_doc = {0}", id_doc), conn);
                SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                // lock есть выше, при вызове этой функции
                using (SqlDataReader dr = SelectCommand.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        List<Ozekidata> lod = new List<Ozekidata>();
                        while (dr.Read() == true)
                        {
                            Ozekidata r = new Ozekidata();
                            r.msgout = new Ozekimessageout();
                            r.id_doc = Convert.ToInt32(dr["id_doc"]);
                            r.id_card = Convert.ToInt32(dr["id_card"]);
                            r.id_branchcard = Convert.ToInt32(dr["id_branchcard"]);
                            if (dr["pan"] != DBNull.Value)
                            {
                                r.last4 = ((string)dr["pan"]).Trim();
                                if (r.last4.Length > 4)
                                    r.last4 = r.last4.Substring((r.last4.Length - 4));
                            }
                            if (dr["phone"] != DBNull.Value) r.msgout.receiver = (string)dr["phone"];
                            if (dr["company"] != DBNull.Value) r.company = ((string)dr["company"]).Trim();
                            if (dr["fio"] != DBNull.Value) r.fio = (string)dr["fio"];
                            if (dr["bin"] != DBNull.Value) r.bin = (string)dr["bin"];
                            if (dr["prefix_ow"] != DBNull.Value) r.prefix_ow = (string)dr["prefix_ow"];
                            if (dr["adress"] != DBNull.Value) r.adress = (string)dr["adress"];

                            r.msgout.msg = "тестовая смс";
                            r.msgout.msgtype = "SMS:TEXT";
                            r.msgout.status = "send";
                            lod.Add(r);
                        }
                        rd = lod.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("newData: " + e.Message);
            }
            
            return rd;
        }


        public void saveData(Ozekidata[] data)
        {
            if (data != null && data.Length > 0)
            {
                SqlCommand cmdCheck = new SqlCommand();
                cmdCheck.Connection = conn;
                cmdCheck.CommandTimeout = conn.ConnectionTimeout;
                cmdCheck.CommandText = "select count(*) from ozekidata where id_card=@card";
                cmdCheck.Parameters.Add("card", SqlDbType.Int);

                SqlCommand cmdAP = new SqlCommand();
                cmdAP.Connection = conn;
                cmdAP.CommandTimeout = conn.ConnectionTimeout;
                SqlCommand cmdAPA = new SqlCommand();
                cmdAPA.Connection = conn;
                cmdAPA.CommandTimeout = conn.ConnectionTimeout;
                
                cmdAP.CommandText = "insert into ozekimessageout " +
                                    "(sender,receiver,msg,senttime,receivedtime,operator,msgtype,reference,status,errormsg) " +
                                    "values(@sender,@receiver,@msg,@senttime,@receivedtime,@operator,@msgtype,@reference,@status,@errormsg) select @@identity as newid";

                cmdAP.Parameters.Add("sender", SqlDbType.VarChar,30);
                cmdAP.Parameters.Add("receiver", SqlDbType.VarChar, 30);
                cmdAP.Parameters.Add("msg", SqlDbType.VarChar,1000);
                cmdAP.Parameters.Add("senttime", SqlDbType.DateTime);
                cmdAP.Parameters.Add("receivedtime", SqlDbType.DateTime);
                cmdAP.Parameters.Add("operator", SqlDbType.VarChar, 1000);
                cmdAP.Parameters.Add("msgtype", SqlDbType.VarChar,30);
                cmdAP.Parameters.Add("reference", SqlDbType.VarChar, 30);
                cmdAP.Parameters.Add("status", SqlDbType.VarChar, 30);
                cmdAP.Parameters.Add("errormsg", SqlDbType.VarChar, 250);

                cmdAPA.CommandText = "insert into ozekidata ([id_doc],[id_card],[id_ozeki]) values (@id_doc,@id_card,@id_ozeki)  select @@identity as newid";

                cmdAPA.Parameters.Add("id_doc", SqlDbType.Int);
                cmdAPA.Parameters.Add("id_card", SqlDbType.Int);
                cmdAPA.Parameters.Add("id_ozeki", SqlDbType.Int);

                try
                {

                    for (int i = 0; i < data.Length; i++)
                    {
                        if (data[i].msgout == null) continue;

                        cmdCheck.Parameters["card"].Value = data[i].id_card;
                        object obj = cmdCheck.ExecuteScalar();
                        if (Convert.ToInt32(obj) > 0)
                            continue;

                        cmdAP.Parameters["sender"].Value = (data[i].msgout.sender != null) ? data[i].msgout.sender : Convert.DBNull;
                        cmdAP.Parameters["receiver"].Value = (data[i].msgout.receiver != null) ? data[i].msgout.receiver : Convert.DBNull;
                        cmdAP.Parameters["msg"].Value = (data[i].msgout.msg != null) ? data[i].msgout.msg : Convert.DBNull;
                        cmdAP.Parameters["senttime"].Value = (data[i].msgout.senttime != DateTime.MinValue) ? data[i].msgout.senttime : Convert.DBNull;
                        cmdAP.Parameters["receivedtime"].Value = (data[i].msgout.receivedtime != DateTime.MinValue) ? data[i].msgout.receivedtime : Convert.DBNull;
                        cmdAP.Parameters["operator"].Value = (data[i].msgout.operators != null) ? data[i].msgout.operators : Convert.DBNull; 
                        cmdAP.Parameters["msgtype"].Value = (data[i].msgout.msgtype != null) ? data[i].msgout.msgtype : Convert.DBNull;
                        cmdAP.Parameters["reference"].Value = (data[i].msgout.reference != null) ? data[i].msgout.reference : Convert.DBNull;
                        cmdAP.Parameters["status"].Value = (data[i].msgout.status != null) ? data[i].msgout.status : Convert.DBNull;
                        cmdAP.Parameters["errormsg"].Value = (data[i].msgout.errormsg != null) ? data[i].msgout.errormsg : Convert.DBNull; ; ;
                        
                        obj = cmdAP.ExecuteScalar();
                        
                        data[i].msgout.id = Convert.ToInt32(obj);
                        
                        cmdAPA.Parameters["id_doc"].Value = data[i].id_doc;
                        cmdAPA.Parameters["id_card"].Value = data[i].id_card;
                        cmdAPA.Parameters["id_ozeki"].Value = data[i].msgout.id;
                        
                        obj = cmdAPA.ExecuteScalar();
                        
                        data[i].id = Convert.ToInt32(obj);


                    }
                }
                catch (Exception e)
                {
                    throw new Exception("saveData: " + e.Message);
                }
            }
        }
    }

}
