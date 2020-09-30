using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using OstCard.Data;
using System.Data.SqlClient;

namespace CardPerso
{
    public partial class StorDocCardEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int id_doc = 0; int id_type = 0; int id_branch;
        int id_act = 0;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                return;
            }
            lock (Database.lockObjectDB)
            {
                id_doc = Convert.ToInt32(Request.QueryString["id_doc"]);
                id_act = Convert.ToInt32(Request.QueryString["id_act"]);
                id_type = Convert.ToInt32(Request.QueryString["type_doc"]);
                if (Request.QueryString["id_branch"] != "")
                    id_branch = Convert.ToInt32(Request.QueryString["id_branch"]);

                if (!IsPostBack)
                {
                    RefrFileCard(true);
                }
            }
        }

        private void RefrFileCard(bool sel_file)
        {
            rbType.Enabled = true;
            dListFile.Enabled = true;

            if (id_type == (int)TypeDoc.CardToStorage || id_type == (int)TypeDoc.PersoCard)
            {
                if (sel_file) rbType.SelectedIndex = 0;
                else rbType.SelectedIndex = 1;

                dListFile.Items.Clear();
                ds.Clear();
                res = Database.ExecuteQuery(String.Format("select distinct FileOw from Cards where (id_stat=1) and (id not in (select id_card from V_CardsTypeDocs where type={0}))",id_type.ToString()), ref ds, null);

                dListFile.DataSource = ds.Tables[0];
                dListFile.DataTextField = "FileOw";
                dListFile.DataValueField = "FileOw";
                dListFile.DataBind();

                if (dListFile.Items.Count > 0) dListFile.SelectedIndex = 0;

                dListCard.Items.Clear();
                ds.Clear();
                res = Database.ExecuteQuery(String.Format("select id,pan from Cards where (id_stat=1) and (id not in (select id_card from V_CardsTypeDocs where type={0})) order by pan",id_type.ToString()), ref ds, null);
               
                dListCard.DataSource = ds.Tables[0];
                dListCard.DataTextField = "pan";
                dListCard.DataValueField = "id";
                dListCard.DataBind();

                if (dListCard.Items.Count > 0) dListCard.SelectedIndex = 0;
            }

            if (id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.SendToClient || id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.ReceiveToBank)
            {
                rbType.Enabled = false;
                rbType.SelectedIndex = 1;

                dListFile.Items.Clear();
                dListFile.Enabled = false;

                dListCard.Items.Clear();
                ds.Clear();

                if (id_type == (int)TypeDoc.SendToFilial)
                    res = Database.ExecuteQuery(String.Format("select id,pan from Cards where (id_stat=2) and (id_BranchCard={0}) and ((id in (select id_card from V_CardsTypeDocs where type=10))  or (id not in (select id_card from V_CardsTypeDocs where type=5))) order by pan", id_branch.ToString()), ref ds, null);
                if (id_type == (int)TypeDoc.ReceiveToFilial)
                    res = Database.ExecuteQuery(String.Format("select Cards.id,Cards.pan from Cards_StorageDocs inner join Cards on Cards_StorageDocs.id_card=Cards.id where (Cards_StorageDocs.id_doc={0}) and (Cards.id not in (select id_card from Cards_StorageDocs where id_doc={1})) order by pan", id_act.ToString(),id_doc.ToString()), ref ds, null);
              //  if (id_type == 7)
              //      res = Database.ExecuteQuery(String.Format("select id,pan from Cards where (id_stat=4) and (id_BranchCard={0}) and (id not in (select id_card from V_CardsTypeDocs where type=7)) order by pan", id_branch.ToString()), ref ds, null);
                if (id_type == (int)TypeDoc.SendToBank)
                    res = Database.ExecuteQuery(String.Format("select id,pan from V_Cards where (id_stat=4) and ((id_BranchCard={0}) or (id_BranchCard_parent={0})) and (id not in (select id_card from V_CardsTypeDocs where type=9)) order by pan", id_branch.ToString()), ref ds, null);
                if (id_type == (int)TypeDoc.ReceiveToBank)
                    //res = Database.ExecuteQuery(String.Format("select id,pan from Cards where (id_stat=5) and (id_BranchCard={0}) and (id not in (select id_card from V_CardsTypeDocs where type={1})) order by pan", id_branch.ToString(), id_type.ToString()), ref ds, null);
                    res = Database.ExecuteQuery(String.Format("select Cards.id,Cards.pan from Cards_StorageDocs inner join Cards on Cards_StorageDocs.id_card=Cards.id where (Cards_StorageDocs.id_doc={0}) and (Cards.id not in (select id_card from Cards_StorageDocs where id_doc={1})) order by pan", id_act.ToString(), id_doc.ToString()), ref ds, null);
 
                dListCard.DataSource = ds.Tables[0];
                dListCard.DataTextField = "pan";
                dListCard.DataValueField = "id";
                dListCard.DataBind();

                if (dListCard.Items.Count > 0) dListCard.SelectedIndex = 0;
            }
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                return;
            }
            lock (Database.lockObjectDB)
            {
                SqlCommand sqCom = Database.Conn.CreateCommand();

                sqCom.CommandText = "select priz_gen from StorageDocs where id=" + id_doc.ToString();
                object obj = sqCom.ExecuteScalar();
                if (Convert.ToBoolean(obj))
                {
                    lbInform.Text = "Нельзя добавить карты в подтвержденный документ";
                    return;
                }
                if (rbType.SelectedIndex == 0)
                {
                    if (dListFile.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите исходный файл";
                        dListFile.Focus();
                        return;
                    }
                    WebLog.LogClass.WriteToLog("StorDocCardEdit.bSaveClick InsertCards Start");
                    sqCom.CommandText = "insert into Cards_StorageDocs (id_doc,id_card) select @id_doc,id from Cards where (FileOw=@FileOw) and (id not in (select id_card from V_CardsTypeDocs where type={0}))";
                    sqCom.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    sqCom.Parameters.Add("@FileOw", SqlDbType.VarChar).Value = dListFile.SelectedItem.Value;
                    Database.ExecuteNonQuery(sqCom, null);
                    WebLog.LogClass.WriteToLog("StorDocCardEdit.bSaveClick InsertCards End");
                }
                if (rbType.SelectedIndex == 1)
                {
                    if (dListCard.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите карту";
                        dListCard.Focus();
                        return;
                    }
                    WebLog.LogClass.WriteToLog("StorDocCardEdit.bSaveClick InsertCard  Start id_doc={0}, id_card={1}", id_doc, dListCard.SelectedItem.Value);
                    sqCom.CommandText = "insert into Cards_StorageDocs (id_doc,id_card) values (@id_doc,@id_card)";
                    sqCom.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    sqCom.Parameters.Add("@id_card", SqlDbType.VarChar).Value = dListCard.SelectedItem.Value;
                    Database.ExecuteNonQuery(sqCom, null);
                    WebLog.LogClass.WriteToLog("StorDocCardEdit.bSaveClick InsertCard End id_doc={0}, id_card={1}", id_doc, dListCard.SelectedItem.Value);
                }

                lbInform.Text = "Карта добавлена в документ. Обновление после закрытия формы.";

                if (rbType.SelectedIndex == 0)
                    RefrFileCard(true);
                if (rbType.SelectedIndex == 1)
                    RefrFileCard(false);

            }
        }
    }
}
