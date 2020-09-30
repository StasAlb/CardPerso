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
using System.Web.Configuration;
using System.Security.Cryptography;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Xml.Linq;
using System.Net;
using System.Net.Mail;
using System.IO;
using OstCard.Data;
using System.Data.SqlClient;
using CardPerso.Administration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Newtonsoft.Json.Serialization;
using CardPerso;

namespace CardPerso
{
    [DataContract]
    public class DocumentData
    {
        [DataMember(Order = 1, Name = "documentId")]
        public string documentId;
        [DataMember(Order = 2, Name = "operationType")]
        public string operationType;
        [DataMember(Order = 3, Name = "cardsState")]
        public string cardsState;
        [DataMember(Order = 4, Name = "dateTime")]
        public string dateTime;
        [DataMember(Order = 5, Name = "cards")]
        public List<CardData> cards = new List<CardData>();
        public void Clear()
        {
            documentId = "";
            operationType = "";
            cardsState = "";
            dateTime = "";
            cards.Clear();
        }
    }
    [DataContract]
    public class CardData
    {
        [DataMember(Order = 1, Name = "externalCardId")]
        public string externalCardId;
        [DataMember(Order = 2, Name = "number")]
        public string number;
        [DataMember(Order = 3, Name = "branch")]
        public string branch;
    }
    public partial class StorDoc : System.Web.UI.Page
    {
        private string res="";
        private bool allcards = false;
        private string docTempName = "";
//        private DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();
        SqlConnection conn = null;
//        object FuncClass.LockObject = new object();
        object lockObjectLog = new object();

        int branch_main_filial = 0;
        int current_branch_id = 0;
        bool perso = false;
        bool reconfirm = false;

        int id_branchExpertiza = 0;
        string urlrmk = "";
        private bool usermk = false;
        string depname = "";
        DocumentData documentJson = new DocumentData();

        private ClientType clientType;

        public String getDepName()
        {
            lock (OstCard.Data.Database.lockObjectDB)
            {
                ServiceClass sc = new ServiceClass();
                int idBranch = sc.BranchId(Page.User.Identity.Name);
                if (idBranch < 1) return "";
                DataSet ds = new DataSet();
                OstCard.Data.Database.ExecuteQuery("select ident_dep from branchs where id=" + idBranch.ToString(), ref ds, null);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    return ds.Tables[0].Rows[0]["ident_dep"].ToString();
                return "";
            }
        }

        public bool IsExpertiza()
        {
            if (gvDocs.Rows.Count > 0 && gvDocs.SelectedIndex >= 0)
            {
                int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                if (id_type == (int)TypeDoc.ReceiveToFilialExpertiza || id_type == (int)TypeDoc.SendToExpertiza
                                                                     || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza)
                {
                    //gvCards.Columns[7].HeaderText = "Жалоба клиента";
                    return true;
                }
            }
            //gvCards.Columns[7].HeaderText = "Комментарий";
            return false;
        }

        public string getgvCards() { return "gvCards"; }
        private string userToFilialFilial="";

        protected bool isMainFilial()
        {
            return (branch_main_filial > 0 && branch_main_filial == current_branch_id);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }

            depname = getDepName();

            lock (Database.lockObjectDB)
            {
                if (!sc.UserAction(User.Identity.Name, Restrictions.MovingView))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);
                conn = new SqlConnection();
                conn.ConnectionString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
                conn.Open();
                ClientScript.RegisterHiddenField("resd", "");
                ClientScript.RegisterHiddenField("resdd", "");
                ClientScript.RegisterHiddenField("resde", "");

                perso=sc.UserAction(User.Identity.Name, Restrictions.Perso);
                current_branch_id = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(current_branch_id,perso);
                reconfirm = sc.UserAction(User.Identity.Name, Restrictions.MovingReconfirm);
                urlrmk = ConfigurationSettings.AppSettings["URLRMKService"];
                try
                {
                    usermk = Convert.ToBoolean(ConfigurationSettings.AppSettings["EnableRMK"]);
                }
                catch (Exception exception)
                {
                    usermk = false;
                }
                clientType = ClientType.AkBars;
                try
                {
                    string str = ConfigurationManager.AppSettings["ClientType"];
                    if (str.ToLower() == "uzcard")
                        clientType = ClientType.Uzcard;
                }
                catch
                {
                }
                if (IsPostBack)
                    return;
                txtCardNumber.Text = "";
                txtCardNumberInfo.Text = "";
                txtClientCompilant.Text = "";
                lbSearch.Text = "where (priz_gen=0)";
                Refr(0, true);
                lbViewP.Text = "P";
                ViewTypePanel();
                lbInform.Text = "";
                if (gvDocs.Rows.Count > 0)
                    gvDocs_SelectedIndexChanged(null, null);
                GC.Collect();
            }
        }
        private string ExecuteQuery(string query, ref DataSet dataSet, SqlTransaction trans)
        {
            try
            {
                using (SqlDataAdapter dataAdapter = new SqlDataAdapter())
                {
                    dataAdapter.SelectCommand = new SqlCommand(query, conn);
                    dataAdapter.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    if (trans != null)
                        dataAdapter.SelectCommand.Transaction = trans;
                    dataAdapter.Fill(dataSet);
                }
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLogErr("StorDoc.ExecuteQuery: query={0}, err={1}", query, e.Message);
            }
            return "";
        }
        public string ExecuteScalar(SqlCommand comm, ref object obj, SqlTransaction transaction)
        {
            try
            {
                comm.CommandTimeout = conn.ConnectionTimeout;
                comm.Connection = conn;
                if (transaction != null)
                    comm.Transaction = transaction;
                obj = comm.ExecuteScalar();
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLogErr("StorDoc.ExecuteScalar: commandText={0}, err={1}", comm.CommandText, e.Message);
            }
            return "";
        }
        public string ExecuteScalar(string query, ref object obj, SqlTransaction transaction)
        {
            try
            {
                SqlCommand comm = conn.CreateCommand();
                comm.CommandTimeout = conn.ConnectionTimeout;
                if (transaction != null)
                    comm.Transaction = transaction;
                comm.CommandText = query;
                obj = comm.ExecuteScalar();
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLogErr("StorDoc.ExecuteScalar: query={0}, err={1}", query, e.Message);
            }
            return "";
        }
        private void Refr(int rowindex,bool conf)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            int currentBranchId = Convert.ToInt32(sc.BranchId(User.Identity.Name));

            //gvDocs.DataSource = null;
            //gvDocs.DataBind();

            DataSet ds = new DataSet();
            ds.Clear();
            string branchId = "";
            if (!sc.UserAction(User.Identity.Name, Restrictions.AllData))
            {
                branchId = " id_branch=" + sc.BranchId(User.Identity.Name).ToString();
                if (sc.UserAction(User.Identity.Name, Restrictions.Perso))
                    branchId = " role_tp=2";
                if (sc.UserAction(User.Identity.Name, Restrictions.Filial))
                {
                    branchId = String.Format(" (id_branch={0} or (type=19 and id_act={0}))", sc.BranchId(User.Identity.Name));
                    if (branch_main_filial > 0)
                    {
                        //branchId += " and (case";
                        if (currentBranchId == branch_main_filial)
                        {
                            //branchId += String.Format(" or ((id_branch={0} or id_branch in (select id from Branchs where id_parent={0})) and ( type=10 or type=11))", branch_main_filial);
                        }
                        else
                            branchId += " and (type!=10 and type!=11)";
                    }
                    branchId += " and (case when type=5 then 0 else 1 end)=1";//" and (case when type=5 and id_branch=106 then 0 else 1 end)=1";
                }
                //if (SingleQuery.IsAccountable(User.Identity.Name))
                //{
                //    branchId += $" and (type={(int)TypeDoc.ReceiveToPodotchet} or type={(int)TypeDoc.SendToClientFromPodotchet} or type={(int)TypeDoc.ReturnFromPodotchet}) and " +
                //                $" id in (select id_doc from AccountablePerson_StorageDocs a1 inner join AccountablePerson a2 on a1.id_person=a2.id where UserId={Session["CurrentUserId"]})";
                //}
                //else
                //{
                //    branchId += $" and (type<>{(int)TypeDoc.ReceiveToPodotchet} and type<>{(int)TypeDoc.SendToClientFromPodotchet} and type<>{(int)TypeDoc.ReturnFromPodotchet})";
                //}
                if (lbSearch.Text.Trim().Length == 0)
                    branchId = "where " + branchId;
                else
                    branchId = " and " + branchId;
            }
            else
            {
                // исключение Казанского филиала по операциям SendToBank = 9, ReceiveToBank = 10, DeleteBrak = 11,
                //branchId = String.Format(" ((id_branch!={0} and id_branch not in (select id from Branchs where id_parent={0}) and (type=9 or type=10 or type=11)) or id_branch is null)", 106);
                //if (lbSearch.Text.Trim().Length==0)
                //    branchId = "where " + branchId;
                //else
                //    branchId = " and " + branchId;
            }
            if (lbSort.Text.Trim().Length == 0)
                lbSort.Text = " order by date_doc desc, id desc ";
                        
            res = ExecuteQuery(String.Format("select *,case	when id_act!=0 then ' <-- ' + (select b.department from StorageDocs sd join branchs b on b.id=sd.id_branch where sd.id=vsd.id_act) else NULL end as frombranch, (select top 1 id_parent from Branchs where id=vsd.id_branch) as parentbranch  from V_StorageDocs vsd {0} {2} {1} ", lbSearch.Text,lbSort.Text,branchId), ref ds, null);

            
            ds.Tables[0].Columns.Add("fulldate");
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                //if (i < gvDocs.PageIndex * gvDocs.PageSize || i > (gvDocs.PageIndex + 1) * gvDocs.PageSize)
//                    continue;
                ds.Tables[0].Rows[i]["fulldate"] = String.Format("{0:dd.MM.yyyy} {1}", Convert.ToDateTime(ds.Tables[0].Rows[i]["date_doc"]), ds.Tables[0].Rows[i]["time_doc"].ToString());

                int type = Convert.ToInt32(ds.Tables[0].Rows[i]["type"]);

                int id = Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);
                //~!
                int branch_main = 0;
                int branch_curr_id = 0;
                object o = ds.Tables[0].Rows[i]["id_branch"];
                if (o != null && o != DBNull.Value)
                {
                    //branch_main = BranchStore.getBranchMainFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]), false);
                    if (ds.Tables[0].Rows[i]["parentbranch"] != DBNull.Value)
                        branch_main = Convert.ToInt32(ds.Tables[0].Rows[i]["parentbranch"]);
                    branch_curr_id = Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]);
                    if (branch_main == 0)
                        branch_main = branch_curr_id;
                }
                if (perso && (type == (int)TypeDoc.ReceiveToFilialExpertiza || type == (int)TypeDoc.SendToExpertiza))
                {
                    ds.Tables[0].Rows.RemoveAt(i);
                    i--;
                    continue;
                }
                //~!
                if (type == (int)TypeDoc.FilialFilial)
                {
                    //~!if (perso == true) // удаляем для персоцентра (Рустем 08.07.2015)
                    if (perso == true && branch_curr_id > 0 && branch_curr_id != current_branch_id)
                    {
                        ds.Tables[0].Rows.RemoveAt(i);
                        i--;
                        continue;
                    }
                    int to = Convert.ToInt32(ds.Tables[0].Rows[i]["id_courier"]);
                    object obj = null;
                    string str = ExecuteScalar("select department from branchs where id=" + to.ToString(), ref obj, null);
                    ds.Tables[0].Rows[i]["branch"] += " --> " + obj.ToString();
                    if (ds.Tables[0].Rows[i]["id_act"].ToString() == "1")
                        ds.Tables[0].Rows[i]["type_name"] += " (на уничтожение)";
                }

                if (type == (int)TypeDoc.SendToFilialFilial)
                {
                    if (perso == true) // удаляем для персоцентра (Рустем 08.07.2015)
                    {
                        ds.Tables[0].Rows.RemoveAt(i);
                        i--;
                        continue;
                    }
                    int to = Convert.ToInt32(ds.Tables[0].Rows[i]["id_act"]);
                    if (currentBranchId != 0 && currentBranchId != to)
                    {
                        ds.Tables[0].Rows.RemoveAt(i);
                        i--;
                        continue;
                    }
                    ds.Tables[0].Rows[i]["id_act"] = 0;
                    object obj = null;
                    string str = ExecuteScalar("select department from branchs where id=" + to.ToString(), ref obj, null);
                    ds.Tables[0].Rows[i]["branch"] = obj.ToString() + " --> " + ds.Tables[0].Rows[i]["branch"].ToString();
                }
                if (type == (int)TypeDoc.SendToClientService)
                {
                    if (perso)
                    {
                        ds.Tables[0].Rows.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                // исключение Казанского филиала по операциям SendToBank = 9, ReceiveToBank = 10, DeleteBrak = 11,
                if (type == (int)TypeDoc.SendToBank || type == (int)TypeDoc.ReceiveToBank || type == (int)TypeDoc.DeleteBrak ||
                        type == (int)TypeDoc.SendToClient || type == (int)TypeDoc.ReceiveToFilial
                        || type == (int)TypeDoc.ReceiveFromPodotchet || type == (int)TypeDoc.ReceiveToPodotchet
                        || type == (int)TypeDoc.ReturnFromPodotchet || type == (int)TypeDoc.SendToClientFromPodotchet
                        || type == (int)TypeDoc.SendToPodotchet || type == (int)TypeDoc.WriteOfPodotchet
                        || type == (int)TypeDoc.FromBook124 || type == (int)TypeDoc.GetBook124
                        || type == (int)TypeDoc.ReceiveBook124 || type == (int)TypeDoc.ToBook124
                        || type == (int)TypeDoc.FromGoz || type == (int)TypeDoc.GetGoz
                        || type == (int)TypeDoc.ReceiveGoz || type == (int)TypeDoc.ToGoz
                        || type == (int)TypeDoc.FromGozToPodotchet || type == (int)TypeDoc.ToPodotchetFromGoz
                        || type == (int)TypeDoc.FromPodotchetToGoz || type == (int)TypeDoc.ToGozFromPodotchet
                        || type == (int)TypeDoc.KillingCard)
                {
                    //~!object o=ds.Tables[0].Rows[i]["id_branch"];
                    //~!if (o != null && o != DBNull.Value)
                    //~!{
                    //~!    int branch_main = BranchStore.getBranchMainFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]),false);
                    //-- 19.09.18 
                    //-- добавляю в это условие еще проверку на ReceiveToBank, поскольку Рустем жалуется на отсутствие документов этого типа.
                    //-- возможно нужны еще какие-то типы... да и вообще что тут правильно хз
                    //--if (perso == true && branch_main > 0)
                    if (perso == true && branch_main > 0 && type != (int)TypeDoc.ReceiveToBank)
                    {
                        ds.Tables[0].Rows.RemoveAt(i);
                        i--;
                        continue;
                    }
                    else
                    {
                        if (type == (int)TypeDoc.ReceiveToBank)
                        {
                            if (ds.Tables[0].Rows[i]["frombranch"] != DBNull.Value)
                            {
                                ds.Tables[0].Rows[i]["branch"] = ds.Tables[0].Rows[i]["branch"].ToString() + ds.Tables[0].Rows[i]["frombranch"].ToString();
                            }
                        }
                    }
                    //~!}                    
                }
            }
            gvDocs.DataSource = ds.Tables[0];
            gvDocs.DataBind();
            if (gvDocs.Rows.Count > 0)
            {
                gvDocs.SelectedIndex = rowindex;
                gvDocs.Rows[gvDocs.SelectedIndex].Focus();
                /*
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]),0);
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]),0);
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == 7) //выдача клиенту
                {
                    lbProduct.Visible = false;
                    lbViewP.Text = "C";
                    ViewTypePanel();
                }
                else
                {
                    lbProduct.Visible = true;
                    lbViewP.Text = "P";
                    ViewTypePanel();
                }
                */
                setViewProductCard();
            }
            else
            {
                ViewProducts(-1,0);
                ViewCards(-1,0);
            }

            if (conf)
            {
                lock (Database.lockObjectDB)
                {
                    string s_fld = Database.GetFiledsByUser(sc.UserId(User.Identity.Name), "stordoc", null);
                    if (s_fld != "") FuncClass.HideFields(s_fld, gvDocs);
                }
            }
            SetButtonDoc();
            SetButtonCard();
            //lbCountD.Text = "Документов: " + gvDocs.Rows.Count.ToString();
            lbCountD.Text = "Документов: " + ds?.Tables?[0]?.Rows.Count.ToString();
        }
        private void RefrListDoc()
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            dListDoc.Items.Clear();
            if (gvDocs.Rows.Count > 0)
            {
                int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                //if (id_type != (int)TypeDoc.ReceiveToFilial && id_type != (int)TypeDoc.ReceiveToBank)
                if (id_type != (int)TypeDoc.ReceiveToBank)
                    dListDoc.Items.Add(new ListItem("Текущий документ", gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"].ToString()));
                if (id_type == (int)TypeDoc.PersoCard && FuncClass.ClientType == ClientType.AkBars)
                {
                    dListDoc.Items.Add(new ListItem("Записка на выпуск карт", ((int)TypeFormDoc.OfficeNote).ToString()));
                    dListDoc.Items.Add(new ListItem("Записка на выпуск_1", ((int)TypeFormDoc.OfficeNote1).ToString()));
                    dListDoc.Items.Add(new ListItem("Сорт лист", ((int)TypeFormDoc.SortList).ToString()));
                    dListDoc.Items.Add(new ListItem("2 акта + реестр", ((int)TypeFormDoc.Akt2ReestrPerso).ToString()));
                }
                if (id_type == (int)TypeDoc.PersoCard && FuncClass.ClientType == ClientType.Uzcard)
                {
                    dListDoc.Items.Add(new ListItem("Записка на выпуск карт", ((int)TypeFormDoc.OfficeNote).ToString()));
                }

                if (id_type == (int)TypeDoc.SendToFilial && FuncClass.ClientType == ClientType.AkBars) //для рассылки транспортных карт и для почты россии
                {
//                    if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_deliv"]) > 0)
                        dListDoc.Items.Add(new ListItem("2 акта + реестр", ((int)TypeFormDoc.Akt2Reestr).ToString()));
                    int id = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    object obj = null;
                    string[] strs = ConfigurationSettings.AppSettings["Transport"].Split(',');
                    for (int i = 0; i < strs.Length; i++)
                        strs[i] = String.Format("prefix_ow like '{0}'", strs[i]);
                    ExecuteScalar(String.Format("select count(*) from V_Cards_StorageDocs where id_doc={0} and ({1})", id, String.Join(" or ", strs)), ref obj, null);

                    if ((int)obj > 0)
                        dListDoc.Items.Add(new ListItem("Транспорт: акт", ((int)TypeFormDoc.TransportAct).ToString()));

                    id = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]);
                    ExecuteScalar(String.Format("select count(*) from V_DeliversBranchs where id_branch={0} and deliver_h='{1}'", id, ConfigurationSettings.AppSettings["MailRussia"].ToString()), ref obj, null);
                    if ((int)obj > 0)
                        dListDoc.Items.Add(new ListItem("Почта: акт", ((int)TypeFormDoc.PochtaAct).ToString()));
                    //ExecuteScalar(String.Format("select count(*) from V_DeliversBranchs where id_branch={0} and (deliver_h='{1}' or deliver_h='{2}') ", id, "DHL", "СПСР"), ref obj, null);
                    //if ((int)obj > 0)
                    //{
                    if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_deliv"]) > 0)
                        dListDoc.Items.Add(new ListItem("Сопроводительный документ", ((int)TypeFormDoc.AccompDoc).ToString()));
                    //}
                    strs = ConfigurationManager.AppSettings["CourierActBranch"].Split(',');
                    if (strs.Contains(SingleQuery.BranchName(id)))
                        dListDoc.Items.Add(new ListItem("Курьер: акт", ((int)TypeFormDoc.Courier7777).ToString()));
                }

                if (id_type == (int) TypeDoc.FromBook124 && FuncClass.ClientType == ClientType.AkBars)
                {
                    dListDoc.Items.Add(new ListItem("Ярлык мешка", ((int)TypeFormDoc.Book124Label).ToString()));
                }
                if (dListDoc.Items.Count>0 )
                    dListDoc.SelectedIndex = 0;

                bExcelD.Visible = (dListDoc.Items.Count > 0); 
            }
        }
        private void SetButtonDoc()
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            bEditD.Visible = false;
            bDeleteD.Visible = false;
            bDeleteD2.Visible = false;
            bAutoD.Visible = false;
            bAutoProd.Visible = false;
            bNewProd.Visible = false;
            bNewCard.Visible = false;
            bSostD.Visible = false;
            bExcelD.Visible = false;
            bCardProperty.Visible = false;
            bSostD.Attributes.Remove("OnClick");
            bSostD.Attributes.Add("OnClick", "if(confirm('Изменить состояние?')) { ShowWait('Изменение состояния документа...'); return true; } return false");
            bExcelD.Attributes.Remove("OnClick");
            bChangeDate.Visible = false;
            bChangeDate.Attributes.Remove("OnClick");
            pnlPanSearch.Visible = true;
            if (!sc.UserAction(User.Identity.Name, Restrictions.MovingEdit))
            {
                bNewD.Visible = false;
                bEditProd.Visible = false;
                bDelProd.Visible = false;
                bDelCard.Visible = false;
                return;
            }
          //  gvProducts.Columns[FuncClass.GetFieldIndex("cnt_perso", gvProducts)].Visible = true;
          //  gvProducts.Columns[FuncClass.GetFieldIndex("cnt_brak", gvProducts)].Visible = true;

            if (gvDocs.Rows.Count > 0)
            {

                if (reconfirm == true)
                {
                    bChangeDate.Attributes.Remove("OnClick");
                    bChangeDate.Attributes.Add("OnClick", String.Format("return show_date('dt={0}');", Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]).ToShortDateString()));
                    bChangeDate.Visible = true;
                }
                
                bSostD.Visible = true;
                int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);

                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]) == 1 && reconfirm != true)
                {
                    bSostD.Visible = false;
                }

                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]) != 1)
                {
                    bEditD.Visible = true;
                    bDeleteD.Visible = true;
                    bDeleteD2.Visible = true;
                    bAutoD.Visible = true;
//                    bCardProperty.Visible = true;

                    if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.PinFromStorage 
                        || id_type == (int)TypeDoc.CardToStorage || id_type == (int)TypeDoc.PinToStorage 
                        || id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.PersoCard 
                        || id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.Expendables
                        || id_type == (int)TypeDoc.Reklama) 
                        bNewProd.Visible = true;
                    if (id_type == (int)TypeDoc.CardToStorage || id_type == (int)TypeDoc.SendToFilial 
                        || id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.SendToClient 
                        || id_type == (int)TypeDoc.PersoCard || id_type == (int)TypeDoc.SendToBank 
                        || id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.ReturnToFilial 
                        || id_type == (int)TypeDoc.FilialFilial || id_type == (int)TypeDoc.SendToPodotchet 
                        || id_type == (int)TypeDoc.ReceiveFromPodotchet
                        || id_type == (int)TypeDoc.SendToClientFromPodotchet || id_type == (int)TypeDoc.KillingCard 
                        || id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124
                        || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124
                        || id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz
                        || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz
                        || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                        || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet
                        )
                        bNewCard.Visible = true;
                    if (id_type == (int)TypeDoc.ReceiveToFilialExpertiza)
                    {
                        bNewCard.Visible = true;
                        bEditCard.Visible = true;
                        pnlPanSearch.Visible = false;

                    }
                    if (id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.PersoCard 
                        || id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.ReceiveToFilial
                        || id_type == (int)TypeDoc.FilialFilial
                        || id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124
                        || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124
                        || id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz
                        || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz

                        )
                        bAutoProd.Visible = true;

                    bDeleteD.Attributes.Add("OnClick", String.Format("return confirm('Удалить документ {0} № {1}?');", gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["type_name"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["number_doc"].ToString()));
                    bEditD.Attributes.Add("OnClick", String.Format("return show_stordoc('mode=2&id={0}')", gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id"].ToString()));

                    if (Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]) != DateTime.Now.Date)
                    {
                        bSostD.Attributes.Remove("OnClick");
                        string f = "if (confirm('Изменить состояние?')) { if(show_date('dt=" + Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]).ToShortDateString() + "')) { ShowWait('Изменение состояния документа...'); return true; }} return false;";
                        bSostD.Attributes.Add("OnClick", f);
                        //bSostD.Attributes.Add("OnClick", String.Format("if (confirm('Изменить состояние?')) return show_date('dt={0}'); else return false;", Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]).ToShortDateString()));
                    }

                }
                if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.PinFromStorage)
                {
           //         gvProducts.Columns[FuncClass.GetFieldIndex("cnt_perso", gvProducts)].Visible = false;
           //         gvProducts.Columns[FuncClass.GetFieldIndex("cnt_brak", gvProducts)].Visible = false;
                }

                if (id_type == (int)TypeDoc.SendToFilial)
                {
            //        gvProducts.Columns[FuncClass.GetFieldIndex("cnt_brak", gvProducts)].Visible = false;
                }

                if (id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.ReturnToFilial || id_type == (int)TypeDoc.DeleteBrak || id_type == (int)TypeDoc.ReceiveToBank
                    || id_type == (int)TypeDoc.KillingCard) // по пункту 4 от Рустема
                {
                    bAutoD.Visible = false;
                    bDeleteD2.Visible = false;
                }
                if (id_type == (int)TypeDoc.SendToClient)
                {
                    bAutoD.Visible = false;
                    bDeleteD2.Visible = false;
                    // проверка на то, что есть несколько доверенностей на компанию
                    SqlCommand comm = new SqlCommand();
                    comm.CommandText = "select company from V_Cards_StorageDocs1 where id=@id";
                    comm.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id"]);
                    object obj = null;
                    lock (Database.lockObjectDB)
                    {
                        Database.ExecuteScalar(comm, ref obj, null);
                    }
                    if (obj != null)
                    {
                        comm.Parameters.Clear();
                        //comm.CommandText = "select count(*) from V_Org where embosstitle=@etitle and person is not null";
                        if(branch_main_filial>0)
                            comm.CommandText = "select count(*) from V_Org where embosstitle=@etitle and person is not null and BranchMainFilialId=" + branch_main_filial.ToString();
                        else
                            comm.CommandText = "select count(*) from V_Org where embosstitle=@etitle and person is not null and BranchMainFilialId is null";
                        comm.Parameters.Add("@etitle", SqlDbType.NChar, 50).Value = obj.ToString();
                        string comp = obj.ToString();
                        lock (Database.lockObjectDB)
                        {
                            Database.ExecuteScalar(comm, ref obj, null);
                        }
                        int cnt = Convert.ToInt32(obj);
                        if (cnt > 1)
                            bExcelD.Attributes.Add("OnClick", String.Format("return show_operson('etitle={0}')", comp));
                    }
                }
                //if (id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet
                //                                            || id_type == (int)TypeDoc.SendToClientFromPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet
                //                                            || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet)
                //{
                //    bAutoD.Visible = false;
                //    bDeleteD2.Visible = false;
                //    if (id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet)
                //    {
                //        bNewCard.Visible = false;
                //        pnlPanSearch.Visible = false;
                //        bDelCard.Visible = false;
                //        bEditCard.Visible = false;
                //        if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
                //        {
                //            //bCardProperty.Visible = true;
                //            bDelCard.Visible = true;
                //        }
                //    }
                //}
                //bExcelD.Visible = true;
                bNewProd.Attributes.Add("OnClick", String.Format("return show_stordoc_product('mode=1&id_doc={0}&type_doc={1}','1')", gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["type"].ToString()));
                if (id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.SendToClient || 
                    id_type == (int)TypeDoc.ReturnToFilial || id_type == (int)TypeDoc.FilialFilial || 
                    id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveFromPodotchet ||
                    id_type == (int)TypeDoc.SendToClientFromPodotchet ||
                    id_type == (int)TypeDoc.KillingCard
                    || id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124
                    || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124
                    || id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz
                    || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz
                    || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                    || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet
                    )
                    bNewCard.Attributes.Add("OnClick", String.Format("return show_stordoc_cardrec('id_doc={0}&type_doc={1}&id_branch={2}')", gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["type"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id_branch"].ToString()));
                else
                    if (id_type == (int)TypeDoc.ReceiveToFilialExpertiza)
                    {
                        bNewCard.Attributes.Add("OnClick", "return clickShowCardNumber();");
                        
                    }
                    else
                        bNewCard.Attributes.Add("OnClick", String.Format("return show_stordoc_card('id_doc={0}&type_doc={1}&id_branch={2}&id_act={3}','1')", gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["type"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id_branch"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id_act"].ToString()));
                if (id_type == (int)TypeDoc.SendToClientService)
                {
                    bEditD.Visible = false;
                    bDeleteD.Visible = false;
                    bDeleteD2.Visible = false;
                    bSostD.Visible = false;
                    bAutoD.Visible = false;
                    bChangeDate.Visible = false;
                }
            }
            if (IsExpertiza())
            {
                bDeleteD2.Visible = false;
                bAutoD.Visible = false;
            }

            RefrListDoc();
            bResetFilter.Visible = (lbSearch.Text != "where (priz_gen=0)");
        }
        private void SetButtonProd()
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            bEditProd.Visible = false;
            bDelProd.Visible = false;

            if (gvProducts.Rows.Count > 0)
            {
                int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]) != 1)
                {
                    bEditProd.Visible = true;
                    bDelProd.Visible = true;
                    bEditProd.Attributes.Add("OnClick", String.Format("return show_stordoc_product('mode=2&id_doc={0}&type_doc={1}&id_prod={2}','2')", gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["id"].ToString(), gvDocs.DataKeys[Convert.ToInt32(gvDocs.SelectedIndex)].Values["type"].ToString(), gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id"].ToString()));
                    bDelProd.Attributes.Add("OnClick", String.Format("return confirm('Удалить продукт {0} ({1})?');", gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["prod_name"].ToString(), gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["bank_name"].ToString()));
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.DontReceiveToFilial)
                {
                    bNewProd.Visible = false;
                    bEditProd.Visible = false;
                    bDelProd.Visible = false;
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ReceiveToFilialPacket)
                {
                    bNewProd.Visible = false;
                    bEditProd.Visible = false;
                    bDelProd.Visible = false;
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.FromWrapping)
                {
                    bEditProd.Visible = false;
                }
                if (id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.DeleteBrak || id_type == (int)TypeDoc.ReceiveToBank ||
                    id_type == (int)TypeDoc.KillingCard) // По желанию Рустема
                {
                    bEditProd.Visible = false;
                    bDelProd.Visible = false;
                }
                if (IsExpertiza())
                {
                    bEditProd.Visible = false;
                    bDelProd.Visible = false;
                }
            }
        }
        private void SetButtonCard()
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            bEditCard.Visible = false;
            bDelCard.Visible = false;
            bCardProperty.Attributes.Add("OnClick", "");
            bExpertiza.Visible = false;

            if (gvCards.Rows.Count > 0)
            {
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]) != 1)
                {
                    int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.KillingCard) 
                        bEditCard.Visible = true;
                    if (IsExpertiza() == false)
                    {
                        bEditCard.Attributes.Add("OnClick",
                            String.Format("return show_addcomment('id_card={0}')",
                                gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id"].ToString()));
                        bEditCard.ToolTip = "Изменить комментарий";
                    }
                    else
                    {
                        if (gvCards.SelectedIndex >= 0)
                        {
                            txtClientCompilantView.Text = txtClientCompilantEdit.Text = gvCards.Rows[gvCards.SelectedIndex].Cells[7].Text; ///!!!!!!!!!!!!! 7 - comment
                            SqlCommand comm = new SqlCommand();
                            comm.CommandText = string.Format("select resultexpertiza from Cards_StorageDocs where id_doc={0} and id_card={1}", Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"]);
                            Object obj = null;
                            ExecuteScalar(comm, ref obj, null);
                            if (obj != DBNull.Value) txtClientCompilantResult.Text = obj.ToString();
                            else txtClientCompilantResult.Text = "";
                            lCardsProperty.Items.Clear();
                            DataSet ds = new DataSet();
                            //ограничиваем варианты экспертизы двумя результатами: ок и брак карты
                            string res = ExecuteQuery("select id,prop from CardProperty where id=1 or id=6", ref ds, null);
                            if (ds != null && ds.Tables.Count > 0)
                            {
                                lCardsProperty.DataSource = ds.Tables[0];
                                lCardsProperty.DataTextField = "prop";
                                lCardsProperty.DataValueField = "id";
                                lCardsProperty.DataBind();
                                ds.Clear();
                                comm.CommandText = "select id_prop from cards where id=" + gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString();
                                ExecuteScalar(comm, ref obj, null);
                                if (obj != DBNull.Value)
                                {
                                    lCardsProperty.SelectedValue = Convert.ToString(obj);
                                }
                            }
                            bEditCard.Attributes.Add("OnClick", "return clickShowClientCompilant();");

                        }
                        bEditCard.ToolTip = "Изменить жалобу клиента";

                    }
                    if (id_type != (int)TypeDoc.Expertiza)
                        bCardProperty.Attributes.Add("OnClick", String.Format("return show_cardproperty('id_card={0}')", gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString()));
                    else
                        bCardProperty.Attributes.Add("OnClick", "return clickShowCardProperty(0);");
                    bDelCard.Visible = true;
                    //12.12.2019  изменили работу с подотчетниками, теперь подотчетники карты сами не принимают и не возвращают
                    //if (id_type != (int) TypeDoc.ReceiveToPodotchet && id_type != (int) TypeDoc.ReturnFromPodotchet &&
                    //    id_type != (int) TypeDoc.WriteOfPodotchet && id_type != (int) TypeDoc.ReceiveFromPodotchet)
                    //{
                    //    bDelCard.Visible = true;
                    //    bDelCard.Attributes.Add("OnClick",
                    //        String.Format("return confirm('Удалить карту {0}?');",
                    //            gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["pan"].ToString()
                    //                .Trim()));
                    //}
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.SendToBank && gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"] != DBNull.Value && Convert.ToInt32(gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"]) == 4)
                {
                    bCardProperty.Attributes.Add("OnClick", String.Format("return show_cardproperty('id_card={0}&tp=1&id_doc={1}')", gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString(), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"])));
                    bCardProperty.Visible = true;
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ReceiveToBank && gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"] != DBNull.Value && Convert.ToInt32(gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"]) == (int)CardStatus.CourierBank)
                {
                    bCardProperty.Attributes.Add("OnClick", String.Format("return show_cardproperty('id_card={0}&tp=1&id_doc={1}')", gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString(), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"])));
                    bCardProperty.Visible = true;
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.KillingCard && gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"] != DBNull.Value && Convert.ToInt32(gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"]) == 4)
                {
                    bCardProperty.Attributes.Add("OnClick", String.Format("return show_cardproperty('id_card={0}&tp=1&id_doc={1}')", gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString(), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"])));
                    bCardProperty.Visible = true;
                }

                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.SendToClient && gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"] != DBNull.Value && Convert.ToInt32(gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"]) == 4)
                {
                    bCardProperty.Attributes.Add("OnClick", String.Format("return show_cardproperty('id_card={0}&tp=2')", gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString()));
                    bCardProperty.Visible = true;
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.SendToPodotchet && gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"] != DBNull.Value && Convert.ToInt32(gvCards.DataKeys[gvCards.SelectedIndex].Values["id_stat"]) == 4)
                {
                    bCardProperty.Attributes.Add("OnClick", String.Format("return show_cardproperty('id_card={0}&tp=2')", gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString()));
                    bCardProperty.Visible = true;
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.DontReceiveToFilial)
                {
                    bNewCard.Visible = false;
                    bDelCard.Visible = false;
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]) != 1)
                {
                    int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    //11.10.18 - раньше видимо было для родительского филиала не давать добавлять карты для возврата
                    //if (id_type == (int)TypeDoc.SendToBank && isMainFilial() == true)
                    //{
                    //    bNewCard.Visible = false;
                    //    bDelCard.Visible = false;
                    //    bEditCard.Visible = false;
                    //}
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.SendToClientService)
                {
                    bNewCard.Visible = false;
                    bDelCard.Visible = false;
                    bEditCard.Visible = false;
                    bCardProperty.Visible = false;
                }
                if (IsExpertiza())
                {
                    bExpertiza.Visible = true;

                    if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]) == 1)
                    {
                        bNewCard.Visible = false;
                        bDelCard.Visible = false;
                        bEditCard.Visible = false;
                        if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.Expertiza)
                        {
                            bCardProperty.Visible = true;
                            bCardProperty.Attributes.Remove("OnClick");
                            bCardProperty.Attributes.Add("OnClick", "return clickShowCardProperty(1);");
                            lCardsProperty.Enabled = false;
                            txtClientCompilantResult.ReadOnly = true;
                        }
                    }
                    else
                    {
                        if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ReceiveToFilialExpertiza)
                        {
                            bNewCard.Visible = true;
                            bDelCard.Visible = true;
                            bEditCard.Visible = true;
                        }
                        else
                        {
                            bNewCard.Visible = false;
                            bDelCard.Visible = false;
                            bEditCard.Visible = false;
                        }
                        if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.Expertiza)
                        {
                            bCardProperty.Visible = true;
                            bCardProperty.Attributes.Remove("OnClick");
                            bCardProperty.Attributes.Add("OnClick", "return clickShowCardProperty(0);");
                            lCardsProperty.Enabled = true;
                            txtClientCompilantResult.ReadOnly = false;
                        }
                    }
                    pnlPanSearch.Visible = false;
                }
            }
        }
        private void ViewProducts(int id,int rowindex)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            DataSet ds = new DataSet();
            ds.Clear();
            res = ExecuteQuery(String.Format("select * from V_Products_StorageDocs where id_doc={0} order by id_sort",id), ref ds, null);
            
            // Пин конверты по картам 04.12.15 (Рустем)
            /*
            Object o = null;
            ExecuteScalar(String.Format("select count(*) from Cards_StorageDocs join Cards on Cards.id=Cards_StorageDocs.id_card  where ispin!=0 and id_doc={0}", id), ref o, null);
            */ 
            if (/*o != null && Convert.ToInt32(o) > 0 && */ds.Tables.Count>0 && ds.Tables[0].Rows.Count>0)
            {
                //int pincount = Convert.ToInt32(o);
            

                // Итог по картам 09.12.15 (Рустем)
                DataRow dr = ds.Tables[0].NewRow();
                dr["prod_name"] = "Итого по картам:";
                dr["cnt_new"] = 0;
                dr["cnt_perso"] = 0;
                dr["cnt_brak"] = 0;
                dr["id_type"] = -1;

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    //if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 2 && Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]) != pincount)
                    //{
                        //???Зачем это??? ds.Tables[0].Rows[i]["cnt_perso"] = pincount;
                    //}
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 1)
                    {
                        dr["cnt_new"] = Convert.ToInt32(dr["cnt_new"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        dr["cnt_perso"] = Convert.ToInt32(dr["cnt_perso"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        dr["cnt_brak"] = Convert.ToInt32(dr["cnt_brak"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                    }
                }
                ds.Tables[0].Rows.Add(dr);
            }
            //....
                  
            
            
            gvProducts.DataSource = ds.Tables[0];
            gvProducts.DataBind();

            if (gvProducts.Rows.Count > 0)
            {
                gvProducts.SelectedIndex = rowindex;
                gvProducts.Rows[gvProducts.SelectedIndex].Focus();
            }

            SetButtonProd();

            int countP = gvProducts.Rows.Count;
            if (countP > 0) countP--;
            lbCountProd.Text = "Продукция: " + countP.ToString();
        }
        private void ViewCards(int id,int rowindex)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            txtCardNumber.Text = "";
            txtCardNumberInfo.Text = "";
            txtClientCompilant.Text = "";
            DataSet ds = new DataSet();
            ds.Clear();
            object obj = null;
            ExecuteScalar(String.Format("select count(*) from Cards_StorageDocs where id_doc={0}", id), ref obj, null);
            int cnt = Convert.ToInt32(obj);
            if (cnt > 10 && !allcards)
            {
                //res = ExecuteQuery(String.Format("select top 10 * from V_Cards_StorageDocs where id_doc={0}", id), ref ds, null);
                res = ExecuteQuery(String.Format("select top 10 *,cards.passport from V_Cards_StorageDocs join cards on cards.id=V_Cards_StorageDocs.id_card where id_doc={0}", id), ref ds, null);
                DataRow dr = ds.Tables[0].NewRow();
                dr["branchcard"] = "И еще " + (cnt - 10).ToString() + " карт";
                dr["fio"] = ""; dr["isPin"] = false;
                ds.Tables[0].Rows.Add(dr);
            }
            else
                //res = ExecuteQuery(String.Format("select * from V_Cards_StorageDocs where id_doc={0}", id), ref ds, null);
                res = ExecuteQuery(String.Format("select *,cards.passport from V_Cards_StorageDocs join cards on cards.id=V_Cards_StorageDocs.id_card where id_doc={0}", id), ref ds, null);
            
            
            gvCards.DataSource = ds.Tables[0];
            gvCards.DataBind();

            if (gvCards.Rows.Count > 0)
            {
                gvCards.SelectedIndex = rowindex;
                gvCards.Rows[gvCards.SelectedIndex].Focus();
            }

            SetButtonCard();

            lbCountCard.Text = "Карты: " + cnt.ToString();
        }
        private void ViewTypePanel()
        {
            if (lbViewP.Text == "P")
            {
                pActionProd.Visible = true;
                pActionCard.Visible = false;
                mpProducts.Visible = true;
                mpCards.Visible = false;
            }

            if (lbViewP.Text == "C")
            {
                pActionProd.Visible = false;
                pActionCard.Visible = true;
                mpCards.Visible = true;
                mpProducts.Visible = false;
            }
        }

        protected void setViewProductCard()
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            lock (Database.lockObjectDB)
            {
                allcards = false;
                int tp = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                switch (tp)
                {
                    // по пункту 4 от Рустема
                    case 9: // возврат ценностей
                        lbProduct.Visible = true;
                        lbCard.Visible = true;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case 12: // возврат в филиал
                    case (7):    //выдача клиенту
                    case (30): //уничтожение ценностей в филиале
                    case (1004): //выдача карт по сервису
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case (14):   //выдача расходных материалов
                        lbCard.Visible = false;
                        lbViewP.Text = "P";
                        ViewTypePanel();
                        break;
                    case (15): //выдача карт для рекламных целей
                        lbCard.Visible = false;
                        lbViewP.Text = "P";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.ReceiveToFilialExpertiza):
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.SendToExpertiza):
                        lbProduct.Visible = true;
                        lbCard.Visible = true;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.ReceiveToExpertiza):
                        lbProduct.Visible = true;
                        lbCard.Visible = true;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.Expertiza):
                        lbProduct.Visible = true;
                        lbCard.Visible = true;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.SendToPodotchet):    //выдача подотчет
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.ReceiveToPodotchet):    //прием подотчет
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.SendToClientFromPodotchet):
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.ReturnFromPodotchet):
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.ReceiveFromPodotchet):
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.WriteOfPodotchet):
                        lbProduct.Visible = false;
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.ToBook124):
                    case ((int)TypeDoc.FromBook124):
                    case ((int)TypeDoc.GetBook124):
                    case ((int)TypeDoc.ReceiveBook124):
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.ToGoz):
                    case ((int)TypeDoc.FromGoz):
                    case ((int)TypeDoc.GetGoz):
                    case ((int)TypeDoc.ReceiveGoz):
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    case ((int)TypeDoc.FromGozToPodotchet):
                    case ((int)TypeDoc.ToPodotchetFromGoz):
                    case ((int)TypeDoc.FromPodotchetToGoz):
                    case ((int)TypeDoc.ToGozFromPodotchet):
                        lbViewP.Text = "C";
                        ViewTypePanel();
                        break;
                    default:
                        lbProduct.Visible = true;
                        lbCard.Visible = true;
                        lbViewP.Text = "P";
                        ViewTypePanel();
                        break;
                }
                gvDocs.Rows[gvDocs.SelectedIndex].Focus();
                int k = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                SetButtonDoc();
                SetButtonCard();
                lbInform.Text = "";
                gvDocs.Rows[gvDocs.SelectedIndex].Focus();
            }
        }
        
        protected void gvDocs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            setViewProductCard();
        }
        protected void bSetFilter_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            lock (Database.lockObjectDB)
            {
                lbSearch.Text = Request.Form["resd"].Replace("[", "'");
                lbSearch.Text = lbSearch.Text.Replace("]", "'");
                Refr(0, false);
            }
        }
        protected void bResetFilter_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            lock (Database.lockObjectDB)
            {
                lbInform.Text = "";
                lbSearch.Text = "where (priz_gen=0)";
                lbSort.Text = "";
                Refr(0, false);
            }
        }
        protected void gvDocs_Sorting(object sender, GridViewSortEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            lock (Database.lockObjectDB)
            {
                int i = FuncClass.GetSortIndex(e.SortExpression, gvDocs);

                if (lbSortIndex.Text != "")
                {
                    int ind = Convert.ToInt32(lbSortIndex.Text);
                    gvDocs.Columns[ind].HeaderText = gvDocs.Columns[ind].HeaderText.Replace("^", "");
                    gvDocs.Columns[ind].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 153);
                }

                lbSortIndex.Text = i.ToString();

                gvDocs.Columns[i].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(102, 153, 153);

                if ("order by " + e.SortExpression + " asc" == lbSort.Text)
                {
                    lbSort.Text = "order by " + e.SortExpression + " desc";
                    gvDocs.Columns[i].HeaderText = gvDocs.Columns[i].HeaderText + "^";
                }
                else
                    lbSort.Text = "order by " + e.SortExpression + " asc";

                Refr(0, false);
            }
        }
        protected void bNewD_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            lock (Database.lockObjectDB)
            {
                Refr(0, false);
            }
        }
        private void AddStorageDoc()
        {
            int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
            int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
            int id_branch = -1;
            if (gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"].ToString() != "")
                id_branch = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]);
            int id_act = -1;
            if (gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_act"].ToString() != "0")
                id_act = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_act"]);

            AddAutoStorageDoc(id_doc, id_type, id_branch,id_act, null);

            ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
            ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
            lbInform.Text = "";
        }
        private bool CheckProducts(int id_doc, int id_type, int sost)
        {
            int cnt = 0;
            int id_prb = 0;
           
            //ds.Clear();
            //не давать подтверждать пустой документ
            if (sost == 0)
            {
                int cprod = 0;
                int ccard = CntCardsInStorageDocs(id_doc,null);
                int cpers = CntProductsInStorageDocs(id_doc, "perso",0,null);
                int cbrak = CntProductsInStorageDocs(id_doc, "brak",0,null);
                if (id_type == 14) // расходники
                    cprod = CntProductsInStorageDocs(id_doc, "new", 2, null);
                if (id_type == 15) // на рекламу
                    cprod = CntProductsInStorageDocs(id_doc, "new", 0, null);
                if (cprod + ccard + cpers + cbrak == 0)
                {
                    lbInform.Text = "Пустой документ подтвердить нельзя";
                    return false;
                }
            }
            //ds.Clear();
            DataSet ds = new DataSet();
            res = ExecuteQuery(String.Format("select * from V_Products_StorageDocs where id_doc={0}",id_doc.ToString()), ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                // не подтверждена
                if (sost == 0)
                {
                    if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.PinFromStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        if (cnt > CntStorage(id_prb, "new",null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")";
                            return false;
                        }
                    }
                    if (id_type == (int)TypeDoc.CardToStorage || id_type == (int)TypeDoc.PinToStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]) 
                            + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"])
                            + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);

                        if (cnt > CntStorage(id_prb, "wrk",null))
                        {
                            lbInform.Text = "На руках нет достаточного кол-ва " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")";
                            return false;
                        }
                    }

                    if (id_type == (int)TypeDoc.SendToFilial)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        if (cnt > CntStorage(id_prb, "new", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва новых " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")"; ;
                            return false;
                        }
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        if (cnt > CntStorage(id_prb, "perso", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва персонализированных " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")"; ;
                            return false;
                        }
                    }

                    if (id_type == (int)TypeDoc.PersoCard)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"])
                            + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);

                        if (cnt > CntStorage(id_prb, "new", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")";
                            return false;
                        }
                    }

                    if (id_type == (int)TypeDoc.DeleteBrak && isMainFilial()==false)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);

                        if (cnt > CntStorage(id_prb, "brak", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")";
                            return false;
                        }
                    }
                }
                //подтверждена
                if (sost == 1)
                {
                    if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.PinFromStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        if (cnt > CntStorage(id_prb, "wrk", null))
                        {
                            lbInform.Text = "На руках нет достаточного кол-ва " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")";
                            return false;
                        }
                    }

                    if (id_type == (int)TypeDoc.CardToStorage || id_type == (int)TypeDoc.PinToStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        if (cnt > CntStorage(id_prb, "new", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва новых " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")"; ;
                            return false;
                        }
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        if (cnt > CntStorage(id_prb, "perso", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва персонализированных " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")"; ;
                            return false;
                        }
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        if (cnt > CntStorage(id_prb, "brak", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва бракованных " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")"; ;
                            return false;
                        }
                    }

                    if (id_type == (int)TypeDoc.PersoCard)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        if (cnt > CntStorage(id_prb, "perso", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва персонализированных " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")"; ;
                            return false;
                        }
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        if (cnt > CntStorage(id_prb, "brak", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва бракованных " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")"; ;
                            return false;
                        }
                    }

                    if (id_type == (int)TypeDoc.ReceiveToBank && isMainFilial() == false)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        if (cnt > CntStorage(id_prb, "brak", null))
                        {
                            lbInform.Text = "В хранилище нет достаточного кол-ва бракованных " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")";
                            return false;
                        }
                    }
                }
            }
            //Возврат карт
            if (id_type == (int)TypeDoc.CardToStorage)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(1, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

                    if (!CheckCntCardsProducts(id_doc, "perso", null))
                    {
                        lbInform.Text = "Количество персонализированных карт не совпадает с количеством заготовок";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(2, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            // Отправка в филиал
            if (id_type == (int)TypeDoc.SendToFilial)
            {
                #region заполняем структуру для json выгрузки
                string[] branchs = (ConfigurationManager.AppSettings.AllKeys.Contains("CourierBranches")) ? ConfigurationManager.AppSettings["CourierBranches"].Split(',') : new string[] {};
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand comm = conn.CreateCommand())
                    {
                        comm.CommandText = $"select Cards.pan, cards.CardIdOW, Branchs.ident_dep from Cards_StorageDocs left join Cards on Cards_StorageDocs.id_card=Cards.id left join Branchs on Cards.id_BranchCard=Branchs.id where Cards_StorageDocs.id_doc={id_doc}";
                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                CardData c = new CardData();
                                c.externalCardId = (dr["CardIdOW"] == DBNull.Value) ? "" : dr["CardIdOW"].ToString();
                                c.number = (dr["pan"] == DBNull.Value) ? "" : dr["pan"].ToString().Trim();
                                c.branch = (dr["ident_dep"] == DBNull.Value) ? "" : dr["ident_dep"].ToString();
                                if (branchs.Contains(c.branch))
                                    documentJson.cards.Add(c);
                            }
                            dr.Close();
                        }
                    }
                    conn.Close();
                }
                #endregion
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(2, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                    if (!CheckCntCardsProducts(id_doc, "perso", null))
                    {
                        lbInform.Text = "Количество отправляемых карт не совпадает с количеством заготовок";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(3, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            // отправка филиал -> филиал
            if (id_type == (int)TypeDoc.SendToFilialFilial)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                    if (!CheckCntCardsProducts(id_doc, "perso", null))
                    {
                        lbInform.Text = "Количество отправляемых карт не совпадает с количеством заготовок";
                        return false;
                    }
                }
                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(3, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            // Прием в филиале поштучно
            if (id_type == (int)TypeDoc.ReceiveToFilial)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(3, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

                    if (!CheckCntCardsProducts(id_doc, "perso", null))
                    {
                        lbInform.Text = "Количество персонализированных карт не совпадает с количеством заготовок";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null) && !CheckStatusCardsInStorageDocs(6, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            //прием в филиале пакетно
            if (id_type == (int)TypeDoc.ReceiveToFilialPacket)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(3, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

                    if (!CheckCntCardsProducts(id_doc, "perso", null))
                    {
                        lbInform.Text = "Количество персонализированных карт не совпадает с количеством заготовок";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(22, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Отказ в прием в филиале - возврат в центр на уничтожение
            if (id_type == (int)TypeDoc.DontReceiveToFilial)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(3, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

                    if (!CheckCntCardsProducts(id_doc, "brak", null))
                    {
                        lbInform.Text = "Количество персонализированных карт не совпадает с количеством заготовок";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(5, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            //Персонализация карт (универсальный)
            if (id_type == (int)TypeDoc.PersoCard)
            {
                #region заполняем структуру для json выгрузки
                string[] branchs = (ConfigurationManager.AppSettings.AllKeys.Contains("CourierBranches")) ? ConfigurationManager.AppSettings["CourierBranches"].Split(',') : new string[] { };
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand comm = conn.CreateCommand())
                    {
                        comm.CommandText = $"select Cards.pan, cards.CardIdOW, Branchs.ident_dep from Cards_StorageDocs left join Cards on Cards_StorageDocs.id_card=Cards.id left join Branchs on Cards.id_BranchCard=Branchs.id where Cards_StorageDocs.id_doc={id_doc}";
                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                CardData c = new CardData();
                                c.externalCardId = (dr["CardIdOW"] == DBNull.Value) ? "" : dr["CardIdOW"].ToString();
                                c.number = (dr["pan"] == DBNull.Value) ? "" : dr["pan"].ToString().Trim();
                                c.branch = (dr["ident_dep"] == DBNull.Value) ? "" : dr["ident_dep"].ToString();
                                if (branchs.Contains(c.branch))
                                    documentJson.cards.Add(c);
                            }
                            dr.Close();
                        }
                    }
                    conn.Close();
                }
                #endregion
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(1, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

                    if (!CheckCntCardsProducts(id_doc, "perso", null))
                    {
                        lbInform.Text = "Количество персонализированных карт не совпадает с количеством заготовок";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocsPerso(id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Выдача клиенту
            if (id_type == (int)TypeDoc.SendToClient)
            {
                if (sost == 0)
                {
                    int ccnt = 0;
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = $"select Count(Cards_StorageDocs.id) as cnt from Cards_StorageDocs " +
                                    "inner join Cards on Cards_StorageDocs.id_card = Cards.id where (Cards.id_stat <> 19 and Cards.id_stat <> 4) " +
                                    "and (Cards_StorageDocs.id_doc=@iddoc)";
                            comm.Parameters.Add("@iddoc", SqlDbType.Int).Value = id_doc;
                            ccnt = Convert.ToInt32(comm.ExecuteScalar());
                        }
                        conn.Close();
                    }
                    //можно выдавать из филиала и из книги 124
                    if (ccnt > 0)
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(8, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            // Прием в филиале на экспертизу
            if (id_type == (int)TypeDoc.ReceiveToFilialExpertiza)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(8, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(11, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Отправка из филиала на экспертизу
            if (id_type == (int)TypeDoc.SendToExpertiza)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(11, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(12, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Прием из филиала на экспертизу
            if (id_type == (int)TypeDoc.ReceiveToExpertiza)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(12, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(13, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            // Прием из филиала на экспертизу
            if (id_type == (int)TypeDoc.Expertiza)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(13, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    bool ccc = false;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        da.SelectCommand = new SqlCommand("", conn);
                        da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                        da.SelectCommand.CommandText = "select Count(Cards_StorageDocs.id) as cnt from Cards_StorageDocs inner join Cards on Cards_StorageDocs.id_card = Cards.id where (Cards.id_stat <> 2 and Cards.id_stat <> 6) and (Cards_StorageDocs.id_doc=@id_doc)";
                        da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                        ccc = (Convert.ToInt32(da.SelectCommand.ExecuteScalar()) == 0);                        
                    }
                    if (!ccc)
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Возврат в филиал
            if (id_type == (int)TypeDoc.ReturnToFilial)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(8, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Возврат в офис
            if (id_type == (int)TypeDoc.SendToBank)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

                    if (!CheckCntCardsProductsReturnToBank(id_doc, null))
                    {
                        lbInform.Text = "Количество карт не совпадает с количеством заготовок";
                        return false;
                    }
                    SqlCommand comm = conn.CreateCommand();
                    comm.CommandText = "select count(*) from V_Cards_StorageDocs where id_doc=@doc and id_prop=1";
                    comm.Parameters.Add("@doc", SqlDbType.Int).Value = id_doc;
                    cnt = Convert.ToInt32(comm.ExecuteScalar());
                    if (Convert.ToInt32(cnt) > 0)
                    {
                        lbInform.Text = "В документе есть карта с состоянием ОК";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(5, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            // филиал - филиал
            if (id_type == (int)TypeDoc.FilialFilial)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(3, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            // уничтожение карт в филиале
            if (id_type == (int) TypeDoc.KillingCard)
            {
                if (sost == 0)
                {
                    //if (!CheckStatusCardsInStorageDocs(4, id_doc, null) && !CheckStatusCardsInStorageDocs(6, id_doc, null))
                    //{
                    //    lbInform.Text = "У одной из карт поменялся статус";
                    //    return false;
                    //}
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = $"select count(*) from V_Cards_StorageDocs where id_doc={id_doc} and id_stat<>4 and id_stat<>6";
                            if (Convert.ToInt32(comm.ExecuteScalar()) > 0)
                            {
                                lbInform.Text = "У одной из карт поменялся статус";
                                conn.Close();
                                return false;
                            }

                            comm.CommandText = $"select count(*) from V_Cards_StorageDocs where id_doc={id_doc} and id_prop=1 and id_stat=4";
                            if (Convert.ToInt32(comm.ExecuteScalar()) > 0)
                            {
                                lbInform.Text = "В документе есть карта с состоянием OK";
                                conn.Close();
                                return false;
                            }
                        }
                        conn.Close();
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(7, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Прием из филиалов
            if (id_type == (int)TypeDoc.ReceiveToBank)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(5, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

                    if (!CheckCntCardsProductsReturnToBank(id_doc, null))
                    {
                        lbInform.Text = "Количество карт не совпадает с количеством заготовок";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocsReturnToBank(id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // уничтожение
            if (id_type == (int)TypeDoc.DeleteBrak)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(6, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }

//                    if (!CheckCntCardsProducts(id_doc, "brak"))
//                    {
//                        lbInform.Text = "Количество уничтожаемых карт не совпадает с количеством заготовок";
//                        return false;
//                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(7, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            #region подотчет
            // Выдача подотчет
            if (id_type == (int)TypeDoc.SendToPodotchet)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    //10.12.2019 раньше статус был филиал подотчет, теперь сразу подотчет
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }

            // Прием подотчет
            if (id_type == (int)TypeDoc.ReceiveToPodotchet)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(14, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            //выдача карт клиенту из под подотчета
            if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(8, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.ReturnFromPodotchet)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(17, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            
            if (id_type == (int)TypeDoc.ReceiveFromPodotchet)
            {
                if (sost == 0)
                {
                    //10.12.2019 раньше статус был подотчет филиал, теперь сразу из подотчет
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.WriteOfPodotchet)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(16, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(8, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            #endregion
            #region книга124
            if (id_type == (int)TypeDoc.ToBook124)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(18, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.GetBook124)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(18, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(19, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.FromBook124)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(19, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(20, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.ReceiveBook124)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(20, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            #endregion
            #region гоз
            if (id_type == (int)TypeDoc.ToGoz)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(23, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.GetGoz)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(23, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(24, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.FromGoz)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(24, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(25, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.ReceiveGoz)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(25, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(4, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            #endregion
            #region гоз-подотчет
            if (id_type == (int)TypeDoc.FromGozToPodotchet)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(24, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.ToPodotchetFromGoz)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(26, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.FromPodotchetToGoz)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }

                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(27, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            if (id_type == (int)TypeDoc.ToGozFromPodotchet)
            {
                if (sost == 0)
                {
                    if (!CheckStatusCardsInStorageDocs(15, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
                if (sost == 1)
                {
                    if (!CheckStatusCardsInStorageDocs(24, id_doc, null))
                    {
                        lbInform.Text = "У одной из карт поменялся статус";
                        return false;
                    }
                }
            }
            #endregion
            return true;
        }
        private void UpdateStorage(int id_doc, int id_type, int sost, DateTime dt, SqlTransaction trans)
        {
            int cnt = 0;
            int id_prb = 0;
            WebLog.LogClass.WriteToLog("StorDoc.UpdateStorage Start id_doc={0}, id_type={1}, sost = {2}, user = {3}, userbranch = {4}", id_doc, id_type, sost, User.Identity.Name, sc.BranchId(User.Identity.Name));                                               
            //ds.Clear();
            DataSet ds = new DataSet();
            res = ExecuteQuery(String.Format("select id,id_prb,cnt_new,cnt_perso,cnt_brak from Products_StorageDocs where id_doc={0}", id_doc.ToString()), ref ds, trans);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                if (sost == 0)
                {
                    if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.PinFromStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageM(cnt, id_prb, "new", trans);
                        StorageP(cnt, id_prb, "wrk", trans);
                    }
                    if (id_type == (int)TypeDoc.CardToStorage || id_type == (int)TypeDoc.PinToStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageM(cnt, id_prb, "wrk", trans);
                        StorageP(cnt, id_prb, "new", trans);

                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        StorageM(cnt, id_prb, "wrk", trans);
                        StorageP(cnt, id_prb, "perso", trans);

                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageM(cnt, id_prb, "wrk", trans);
                        StorageP(cnt, id_prb, "brak", trans);
                    }
                    if (id_type == (int)TypeDoc.SendToFilial)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageM(cnt, id_prb, "new", trans);
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageM(cnt, id_prb, "perso", trans);
                    }
                    if (id_type == (int)TypeDoc.PersoCard)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        StorageM(cnt, id_prb, "new", trans);
                        //StorageP(cnt, id_prb, "perso", trans);

                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageM(cnt, id_prb, "new", trans);
                        StorageP(cnt, id_prb, "brak", trans);
                    }
                    if (id_type == (int)TypeDoc.ReceiveToBank)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageP(cnt, id_prb, "brak", trans);
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageP(cnt, id_prb, "perso", trans);
                    }
                    if (id_type == (int)TypeDoc.ReceiveToExpertiza)
                    {                        
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageP(cnt, id_prb, "perso", trans);
                    }
                    if (id_type == (int)TypeDoc.Expertiza)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        //StorageM(cnt, id_prb, "perso", trans);
                        StorageP(cnt, id_prb, "brak", trans);
                    }
                    if (id_type == (int)TypeDoc.DeleteBrak)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageM(cnt, id_prb, "brak", trans);
                    }
                    if (id_type == (int)TypeDoc.Expendables)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageM(cnt, id_prb, "new", trans);
                    }
                    if (id_type == (int)TypeDoc.Reklama)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageM(cnt, id_prb, "new", trans);
                    }
                    if (id_type == (int)TypeDoc.ToWrapping)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageM(cnt, id_prb, "perso", trans);
                    }
                    if (id_type == (int)TypeDoc.FromWrapping)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageP(cnt, id_prb, "perso", trans);
                    }
                }

                if (sost == 1)
                {
                    if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.PinFromStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageM(cnt, id_prb, "wrk", trans);
                        StorageP(cnt, id_prb, "new", trans);
                    }
                    if (id_type == (int)TypeDoc.CardToStorage || id_type == (int)TypeDoc.PinToStorage)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageM(cnt, id_prb, "new", trans);
                        StorageP(cnt, id_prb, "wrk", trans);

                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        StorageM(cnt, id_prb, "perso", trans);
                        StorageP(cnt, id_prb, "wrk", trans);

                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageM(cnt, id_prb, "brak", trans);
                        StorageP(cnt, id_prb, "wrk", trans);
                    }
                    if (id_type == (int)TypeDoc.SendToFilial)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageP(cnt, id_prb, "new", trans);

                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageP(cnt, id_prb, "perso", trans);
                    }

                    if (id_type == (int)TypeDoc.PersoCard)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageM(cnt, id_prb, "perso", trans);
                        StorageP(cnt, id_prb, "new", trans);

                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageM(cnt, id_prb, "brak", trans);
                        StorageP(cnt, id_prb, "new", trans);
                    }
                    // прием из филиала
                    if (id_type == (int)TypeDoc.ReceiveToBank)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageM(cnt, id_prb, "brak", trans);
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageM(cnt, id_prb, "perso", trans);
                    }
                    if (id_type == (int)TypeDoc.Expertiza)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageM(cnt, id_prb, "brak", trans);
                        //StorageP(cnt, id_prb, "perso", trans);                        
                    }
                    if (id_type == (int)TypeDoc.ReceiveToExpertiza)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageM(cnt, id_prb, "perso", trans);
                    }
                    if (id_type == (int)TypeDoc.DeleteBrak)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);
                        StorageP(cnt, id_prb, "brak", trans);
                    }
                    if (id_type == (int)TypeDoc.Expendables)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageP(cnt, id_prb, "new", trans);
                    }
                    if (id_type == (int)TypeDoc.Reklama)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        StorageP(cnt, id_prb, "new", trans);
                    }
                    if (id_type == (int)TypeDoc.ToWrapping)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageP(cnt, id_prb, "perso", trans);
                    }
                    if (id_type == (int)TypeDoc.FromWrapping)
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        //StorageM(cnt, id_prb, "perso", trans);
                    }
                }
            }
            // операции с картами

            
            //возврат карт
            if (id_type == (int)TypeDoc.CardToStorage)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 2, sost, "DateProd", dt, "", trans);
                if (sost == 1)
                    UpdateCardsDate(id_doc, 1, sost, "DateProd", dt, "", trans);
            }
            // отправка в филилиал
            if (id_type == (int)TypeDoc.SendToFilial)
            {
                if (sost == 0)
                    UpdateCardsCourier(id_doc, 3, sost,dt, gvDocs.DataKeys[gvDocs.SelectedIndex].Values["invoice_courier"].ToString(), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_courier"].ToString(), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["courier_name"].ToString(), trans);
                if (sost == 1)
                    UpdateCardsCourier(id_doc, 2, sost,dt, "", "", "", trans);
            }
            // отправка филиал -> филиал
            if (id_type == (int)TypeDoc.SendToFilialFilial)
            {
                if (sost == 0)
                    UpdateCardsDoc(id_doc, 3, 0, trans);
                if (sost == 1)
                {
                    UpdateCardsDoc(id_doc, 4, 0, trans);
                    // и ставим текущий филиал
                    SqlCommand comm = Database.Conn.CreateCommand();
                    if (trans != null)
                        comm.Transaction = trans;
                    comm.CommandText = "Update Cards set id_branchCurrent=(select id_act from StorageDocs where id=@id_doc) where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    comm.ExecuteNonQuery();
                }
            }
            //прием карт поштучно
            if (id_type == (int)TypeDoc.ReceiveToFilial)
            {
                
                object obj = null; 
                ExecuteScalar($"select id_act from StorageDocs where id={gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_act"]}", ref obj, trans);
                bool isForKilling = (obj == DBNull.Value || (int)obj != 1) ? false : true;
                if (sost == 0)
                {
                    UpdateCardsDate(id_doc, (isForKilling) ? 6 : 4, sost, "DateReceipt", dt, "", null);
                    // 30.10.18 и обнуляем дату выдачи для карт, чтобы после экспертизы можно было выдать карту клиенту
                    SqlCommand comm = Database.Conn.CreateCommand();
                    if (trans != null)
                        comm.Transaction = trans;
                    comm.CommandText = "Update Cards set DateClient=NULL, id_BranchCurrent=@bran where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    comm.Parameters.Add("@bran", SqlDbType.Int).Value = current_branch_id;
                    comm.ExecuteNonQuery();
                }
                if (sost == 1)
                    UpdateCardsDate(id_doc, 3, sost, "DateReceipt", dt, "", trans);
            }
            //прием карт пакетно
            if (id_type == (int)TypeDoc.ReceiveToFilialPacket)
            {
                if (sost == 0)
                {
                    UpdateCardsDate(id_doc, 22, sost, "DateReceipt", dt, "", trans);
                    // 30.10.18 и обнуляем дату выдачи для карт, чтобы после экспертизы можно было выдать карту клиенту
                    SqlCommand comm = Database.Conn.CreateCommand();
                    if (trans != null)
                        comm.Transaction = trans;
                    comm.CommandText = "Update Cards set DateClient=NULL where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    comm.ExecuteNonQuery();
                }
                if (sost == 1)
                    UpdateCardsDate(id_doc, 3, sost, "DateReceipt", dt, "", trans);
            }
            //отказ в приеме карт, отправка обратно в центр на уничтожение
            if (id_type == (int)TypeDoc.DontReceiveToFilial)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 5, sost, "dateSendTerminate", dt, "", false, true, trans);
                if (sost == 1)
                    UpdateCardsDate(id_doc, 3, sost, "dateSendTerminate", dt, "", false, true, trans);
            }
            //универсальный персонализация
            if (id_type == (int)TypeDoc.PersoCard)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 2, sost, "DateProd", dt, "", trans);
                if (sost == 1)
                    UpdateCardsDate(id_doc, 1, sost, "DateProd", dt, "", trans);
            }
            //выдача клиенту
            if (id_type == (int)TypeDoc.SendToClient)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 8, sost, "DateClient",dt,  sc.UserGuid(User.Identity.Name), trans);
                if (sost == 1)
                    UpdateCardsDate(id_doc, 4, sost, "DateClient", dt, "", trans);
            }
            // прием на экспертизу в филиале
            if (id_type == (int)TypeDoc.ReceiveToFilialExpertiza)
            {
                if (sost == 0)
                    UpdateCardsDocExpertiza(id_doc, 11, sost, trans);
                if (sost == 1)
                    UpdateCardsDocExpertiza(id_doc, 8, sost, trans);
            }
            // отправка из филиала на экспертизу
            if (id_type == (int)TypeDoc.SendToExpertiza)
            {
                if (sost == 0)
                    UpdateCardsDocExpertiza(id_doc, 12, sost, trans);
                if (sost == 1)
                    UpdateCardsDocExpertiza(id_doc, 11, sost, trans);
            }
            // прием из филиала на экспертизу
            if (id_type == (int)TypeDoc.ReceiveToExpertiza)
            {
                if (sost == 0)
                    UpdateCardsDocExpertiza(id_doc, 13, sost, trans);
                if (sost == 1)
                    UpdateCardsDocExpertiza(id_doc, 12, sost, trans);
            }
            // экспертиза
            if (id_type == (int)TypeDoc.Expertiza)
            {
                if (sost == 0)
                    UpdateCardsDocExpertiza(id_doc, 130, sost, trans);
                if (sost == 1)
                {
                    UpdateCardsDocExpertiza(id_doc, 13, sost, trans);
                    AddAutoStorageProduct(id_doc, id_type, trans);
                }
            }
            //возврат в филиал
            if (id_type == (int)TypeDoc.ReturnToFilial)
            {
                if (sost == 0)
                    UpdateCardsDoc(id_doc, 4, 1, trans);
                if (sost == 1)
                    UpdateCardsDoc(id_doc, 8, 0, trans);
            }
            //возврат в офис
            if (id_type == (int)TypeDoc.SendToBank)
            {
                if (sost == 0)
                    UpdateCardsTerminate(id_doc, 5, sost, dt, trans);
                if (sost == 1)
                    UpdateCardsTerminate(id_doc, 4, sost, dt, trans);
            }
            //прием из филиалов
            if (id_type == (int)TypeDoc.ReceiveToBank)
            {
                UpdateCardsDateReturnToBank(id_doc, sost, dt,"", trans);
            }
            //уничтожение
            if (id_type == (int)TypeDoc.DeleteBrak)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 7, sost, "DateTerminated", dt,"", trans);
                if (sost == 1)
                    UpdateCardsDate(id_doc, 6, sost, "DateTerminated", dt,"", trans);
            }
            //уничтожение в филиале
            if (id_type == (int)TypeDoc.KillingCard)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 7, sost, "DateTerminated", dt, "", trans);
                if (sost == 1)
                {
                    UpdateCardsDate(id_doc, 4, sost, "DateTerminated", dt, "", trans);
                    SqlCommand comm = Database.Conn.CreateCommand();
                    comm.Transaction = trans;
                    comm.CommandText = $"Update Cards set id_stat=6 where id in (select id_card from V_Cards_StorageDocs where id_doc={id_doc} and id_prop=1)";
                    comm.ExecuteNonQuery();

                }
            }
            //пересылка из филиала в филиал
            if (id_type == (int)TypeDoc.FilialFilial)
            {
                UpdateCardsFilialFilial(id_doc, sost, trans);
            }
            //отправка на упаковку
            if (id_type == (int)TypeDoc.ToWrapping)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 10, sost, "dateToWrapping", dt, "", trans);
                if (sost == 1)
                    UpdateCardsDate(id_doc, 2, sost, "dateToWrapping", dt, "", true, trans);
            }
            //прием с упаковки
            if (id_type == (int)TypeDoc.FromWrapping)
            {
                if (sost == 0)
                    UpdateCardsDate(id_doc, 2, sost, "dateFromWrapping", dt, "", true, trans);
                if (sost == 1)
                    UpdateCardsDate(id_doc, 10, sost, "dateFromWrapping", dt, "", true, trans);
            }
            // выдача подотчет
            if (id_type == (int)TypeDoc.SendToPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 4, sost, trans);
            }

            // прием подотчет
            if (id_type == (int)TypeDoc.ReceiveToPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 14, sost, trans);
            }
            // выдача с подотчета
            if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 8, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans);
            }

            // возврат с подотчета
            if (id_type == (int)TypeDoc.ReturnFromPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 17, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans);
            }
            // прием в хранилище с подотчета
            if (id_type == (int)TypeDoc.ReceiveFromPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 4, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans); // 12.12.2019 теперь возврат идет сразу с подотчета, раньше было с подотчет-филиал
            }
            // списание с подотчета
            if (id_type == (int)TypeDoc.WriteOfPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 8, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 16, sost, trans);
            }
            if (id_type == (int)TypeDoc.ToBook124)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 18, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 4, sost, trans);
            }
            if (id_type == (int)TypeDoc.GetBook124)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 19, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 18, sost, trans);
            }
            if (id_type == (int)TypeDoc.FromBook124)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 20, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 19, sost, trans);
            }
            if (id_type == (int)TypeDoc.ReceiveBook124)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 4, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 20, sost, trans);
            }
            if (id_type == (int)TypeDoc.ToGoz)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 23, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 4, sost, trans);
            }
            if (id_type == (int)TypeDoc.GetGoz)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 24, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 23, sost, trans);
            }
            if (id_type == (int)TypeDoc.FromGoz)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 25, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 24, sost, trans);
            }
            if (id_type == (int)TypeDoc.ReceiveGoz)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 4, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 25, sost, trans);
            }
            if (id_type == (int)TypeDoc.FromGozToPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 24, sost, trans);
            }
            if (id_type == (int)TypeDoc.ToPodotchetFromGoz)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 26, sost, trans);
            }
            if (id_type == (int)TypeDoc.FromPodotchetToGoz)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 27, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans);
            }
            if (id_type == (int)TypeDoc.ToGozFromPodotchet)
            {
                if (sost == 0)
                    UpdateCardsDocPodotchet(id_doc, 24, sost, trans);
                if (sost == 1)
                    UpdateCardsDocPodotchet(id_doc, 15, sost, trans); // 12.12.2019 теперь возврат идет сразу с подотчета, раньше было с подотчет-гоз
            }
            WebLog.LogClass.WriteToLog("StorDoc.UpdateStorage End id_doc={0}, id_type={1}, sost = {2}, user = {3}, userbranch = {4}", id_doc, id_type, sost, User.Identity.Name, sc.BranchId(User.Identity.Name));                                               
        }

        protected void bChangeDate_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            lock (Database.lockObjectDB)
            {
                lbInform.Text = "";
                int id_doc = 0; int id_type = 0; int sost = 0;
                int id_deliv = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_deliv"]);
                id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                SqlCommand comm = Database.Conn.CreateCommand();
                comm.CommandText = "select count(*) from StorageDocs where id="+id_doc.ToString();
                object obj = comm.ExecuteScalar();
                if (Convert.ToInt32(obj) != 1)
                {
                    lbInform.Text = "Данный документ был удален";
                    return;
                }
                DateTime dt = DateTime.MinValue;
                try
                {
                    dt = Convert.ToDateTime(Request.Form["resdd"]);
                }
                catch { }
                id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                sost = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]);
                WebLog.LogClass.WriteToLog("StorDoc.bChangeDate_Click Start id_doc={0}, id_type={1}, sost = {2}, user = {3}, userbranch = {4}", id_doc, id_type, sost, User.Identity.Name, sc.BranchId(User.Identity.Name));                                               
                if (id_deliv == 0)
                {
                    id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    sost = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]);
                    try
                    {
                        ChangeGen(id_doc, sost, dt, null);
                        WebLog.LogClass.WriteToLog(sc.BranchId(User.Identity.Name) + " " + User.Identity.Name + " " + String.Format("Документ смена даты {0} {1} {2}", sc.UserGuid(User.Identity.Name), id_doc, sost));
                    }
                    catch (Exception ex)
                    {
                        WebLog.LogClass.WriteToLog(ex.ToString());
                    }
                }
                else
                {
                    DataSet ds1 = new DataSet();
                    ds1.Clear();
                    res = ExecuteQuery(String.Format("select id,type,priz_gen from StorageDocs where id_deliv={0}", id_deliv.ToString()), ref ds1, null);
                    for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                    {
                            id_doc = Convert.ToInt32(ds1.Tables[0].Rows[i]["id"]);
                            id_type = Convert.ToInt32(ds1.Tables[0].Rows[i]["type"]);
                            sost = Convert.ToInt32(ds1.Tables[0].Rows[i]["priz_gen"]);
                            ChangeGen(id_doc, sost, dt, null);
                    }
                    lbInform.Text = "";
                }
                WebLog.LogClass.WriteToLog("StorDoc.bChangeDate_Click End");
                Refr(0, false);
                
            }
        }


        protected void bSostD_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            }
            int rmko = BranchStore.TypeDocToRMKOperation((TypeDoc)Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]));
            if (usermk && rmko != 0)
            {
                if (sc.UserPassport(User.Identity.Name).Trim().Length < 1)
                {
                    string er = "Для этой операции необходимы паспортные данные пользователя " + User.Identity.Name;
                    ClientScript.RegisterClientScriptBlock(GetType(), "errRMK", "<script type='text/javascript'>$(document).ready(function(){ ShowMessage('" + er + "');});</script>");
                    return;
                }
            }
            lock (Database.lockObjectDB)
            {
                lbInform.Text = "";
                int id_doc = 0; int id_type = 0; int sost = 0; int gen = 0; int number_doc = 0; bool isForKilling = false;
                int id_deliv = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_deliv"]);
                id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                number_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["number_doc"]);
                DateTime datedoc = (DateTime)gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"];
                SqlCommand comm = Database.Conn.CreateCommand();
                comm.CommandText = "select count(*) from StorageDocs where id="+id_doc.ToString();
                object obj = comm.ExecuteScalar();
                if (Convert.ToInt32(obj) != 1)
                {
                    lbInform.Text = "Данный документ был удален";
                    Refr(0, false);
                    return;
                }
                sost = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]);
                id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                if ((TypeDoc)id_type == TypeDoc.FilialFilial)
                {
                    comm.CommandText = $"select id_act from StorageDocs where id={id_doc}";
                    isForKilling = (Convert.ToInt32(comm.ExecuteScalar()) == 1);
                }
                if ((TypeDoc)id_type == TypeDoc.ReceiveToFilial)
                {
                    comm.CommandText = $"select id_act from StorageDocs where id=(select id_act from StorageDocs where id={id_doc})";
                    isForKilling = (Convert.ToInt32(comm.ExecuteScalar()) == 1);
                }
                
                //проверяем, что состоянии документа на странице такое же как в базе (для случаев косячного обновления страницы)
                comm.CommandText = "select priz_gen from StorageDocs where id="+id_doc.ToString();
                int cursost = Convert.ToInt32(comm.ExecuteScalar());
                if (sost != cursost)
                {
                    lbInform.Text = "Состояние уже изменено";
                    Refr(0, false);
                    return;
                }

                DateTime dt = DateTime.MinValue;
                try
                {
                    dt = Convert.ToDateTime(Request.Form["resdd"]);
                }
                catch { }
                id_branchExpertiza = -1;
                if (id_type == (int)TypeDoc.Expertiza && sost == 0)
                {
                    id_branchExpertiza = getIdBranchFromExpertiza();
                    if (id_branchExpertiza == 0)
                    {
                        string er = "Для этой операции необходим код подразделения, где была сдана карта";
                        ClientScript.RegisterClientScriptBlock(GetType(), "errRMK", "<script type='text/javascript'>$(document).ready(function(){ ShowMessage('" + er + "');});</script>");
                        return;
                    }
                }
                WebLog.LogClass.WriteToLog("StorDoc.bSostD_Click Start id_doc={0}, id_type={1}, sost = {2}, user = {3}, userbranch = {4}", id_doc, id_type, sost, User.Identity.Name, sc.BranchId(User.Identity.Name));
                if (id_deliv == 0)
                {
                    id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    sost = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]);
                    //~!if(sost==0 && isMainFilial()==true && id_type == (int)TypeDoc.SendToBank)
                    //~!{
                    //~!    if (false == CheckPropCardsInStorageDocs(10, id_doc, null))
                    //~!    {
                    //~!        lbInform.Text = "Состояние карт должно быть: неверный филиал отправки";
                    //~!        return;
                    //~!    }
                    //~!}
                    if ((id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.ReceiveToFilialExpertiza) && sost != 1)
                    {
                        AddAutoStorageProduct(id_doc, id_type, null);
                    }
                    documentJson.Clear();
                    if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["priz_gen"]) == 0)
                        gen = 1;
                    try
                    {
                        if (!CheckProducts(id_doc, id_type, sost)) return;
                        lbInform.Text = "";
                        #region отсылаем данные в сервис подотчетных лиц
                        try
                        {
                            if (usermk)
                            { 
                                int rmkoper = BranchStore.TypeDocToRMKOperation((TypeDoc)id_type);
                                if ((TypeDoc)id_type == TypeDoc.FilialFilial && isForKilling)
                                    rmkoper = BranchStore.TypeDocToRMKOperation((TypeDoc)id_type, 1);
                                if ((TypeDoc)id_type == TypeDoc.ReceiveToFilial && isForKilling)
                                    rmkoper = BranchStore.TypeDocToRMKOperation((TypeDoc)id_type, 1);
                                if (0 != rmkoper)
                                {
                                    string userPassport = sc.UserPassport(User.Identity.Name);
                                    RMKClient r = new RMKClient(urlrmk);
                                    if (gen != 0)
                                    {
                                        WebLog.LogClass.WriteToLog("Создание RMK Data ({0}) id_doc={1}", urlrmk, id_doc);
                                        BranchStore.clearRMKData(id_doc, conn);
                                        BranchStore bs = new BranchStore(0, "", "");
                                        bs.getBaseProductsFromDocs(id_doc, conn);
                                        int indx = BranchStore.TypeDocToIndexBranchStore((TypeDoc) id_type);
                                        RMKRequestData rqd = new RMKRequestData();
                                        //rqd.operationType = BranchStore.TypeDocToRMKOperation((TypeDoc) id_type).ToString();
                                        rqd.operationType = rmkoper.ToString();
                                        if (dt != DateTime.MinValue) rqd.operDate = dt;
                                        else rqd.operDate = datedoc;
                                        rqd.userSeries = sc.PassportSeries(userPassport);
                                        rqd.userNumber = sc.PassportNumber(userPassport);
                                        rqd.packetId = number_doc.ToString();
                                        rqd.operationId = "";
                                        rqd.branchCode = depname;
                                        if (id_type == (int) TypeDoc.SendToPodotchet)
                                        {
                                            CommonDataBase cdb = new CommonDataBase(null);
                                            int idp = GetPerson(id_doc);
                                            if (idp > 0)
                                            {
                                                AccountablePersonData apd = cdb.getAccountablePersonData(idp);
                                                rqd.accountableSeries = apd.PassportSeries;
                                                rqd.accountableNumber = apd.PassportNumber;
                                            }
                                        }
                                        if (id_type == (int)TypeDoc.ReceiveFromPodotchet)
                                        {
                                            CommonDataBase cdb = new CommonDataBase(null);
                                            int idp = GetPerson(id_doc);
                                            if (idp > 0)
                                            {
                                                // 19.05.2020 меняем их местами
                                                AccountablePersonData apd = cdb.getAccountablePersonData(idp);
                                                //rqd.accountableSeries = rqd.userSeries;
                                                //rqd.accountableNumber = rqd.userNumber;
                                                //rqd.userSeries = apd.PassportSeries;
                                                //rqd.userNumber = apd.PassportNumber;
                                                rqd.accountableSeries = apd.PassportSeries;
                                                rqd.accountableNumber = apd.PassportNumber;
                                            }
                                        }
                                        if (id_type == (int)TypeDoc.ToGozFromPodotchet)
                                        {
                                            // 20.05.2020 креатор теперь гоз, а подотчетник- подотчетник
                                            CommonDataBase cdb = new CommonDataBase(null);
                                            int idp = GetPerson(id_doc);
                                            if (idp > 0)
                                            {
                                                AccountablePersonData apd = cdb.getAccountablePersonData(idp);
                                                //rqd.accountableSeries = rqd.userSeries;
                                                //rqd.accountableNumber = rqd.userNumber;
                                                rqd.accountableSeries = apd.PassportSeries;
                                                rqd.accountableNumber = apd.PassportNumber;
                                            }
                                        }
                                        if (id_type == (int)TypeDoc.FromGozToPodotchet)
                                        {
                                            // 20.05.20202 креатор теперь подотчетник, подотчетник - гоз
                                            // 20.05.20202 (вечер) креатор теперь гоз, подотчетник - подотчетник
                                            CommonDataBase cdb = new CommonDataBase(null);
                                            int idp = GetPerson(id_doc);
                                            if (idp > 0)
                                            {
                                                //rqd.accountableSeries = rqd.userSeries;
                                                //rqd.accountableNumber = rqd.userNumber;
                                                AccountablePersonData apd = cdb.getAccountablePersonData(idp);
                                                rqd.accountableSeries = apd.PassportSeries;
                                                rqd.accountableNumber = apd.PassportNumber;
                                            }
                                        }
                                        if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
                                        {
                                            CommonDataBase cdb = new CommonDataBase(null);
                                            int idp = GetPerson(id_doc);
                                            //19.12.2019 для выдачи делаем и создателя и подчетное лицо одним и тем же
                                            if (idp > 0)
                                            {
                                                AccountablePersonData apd = cdb.getAccountablePersonData(idp);
                                                rqd.userSeries = apd.PassportSeries;
                                                rqd.userNumber = apd.PassportNumber;
                                                rqd.accountableSeries = apd.PassportSeries;
                                                rqd.accountableNumber = apd.PassportNumber;
                                            }
                                        }

                                        if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.ToGoz)
                                        {
                                            ExecuteScalar($"select username from storagedocs inner join aspnet_users on StorageDocs.id_act=aspnet_users.id where storagedocs.id={id_doc}", ref obj, null);
                                            string passport = sc.UserPassport((string)obj);
                                            rqd.accountableSeries = sc.PassportSeries(passport);
                                            rqd.accountableNumber = sc.PassportNumber(passport);
                                        }
                                        if (id_type == (int)TypeDoc.ReceiveBook124)
                                        {
                                            // 25.11.2019 меняем местами того кто принял документ и подотчетное лицо, потому что РМК захотел, чтобы как будто операция шла от имени того, кто сдает карты по 124 книге
                                            rqd.accountableSeries = rqd.userSeries;
                                            rqd.accountableNumber = rqd.userNumber;
                                            ExecuteScalar($"select username from aspnet_users where id=(select user_id from storagedocs where id=(select id_act from storagedocs where id={id_doc}))", ref obj, null);
                                            string passport = sc.UserPassport((string)obj);
                                            rqd.userSeries = sc.PassportSeries(passport);
                                            rqd.userNumber = sc.PassportNumber(passport);
                                            //rqd.accountableSeries = sc.PassportSeries(passport);
                                            //rqd.accountableNumber = sc.PassportNumber(passport);
                                        }
                                        if (id_type == (int)TypeDoc.ReceiveGoz)
                                        {
                                            // 25.11.2019 меняем местами того кто принял документ и подотчетное лицо, потому что РМК захотел, чтобы как будто операция шла от имени того, кто сдает карты по 124 книге
                                            // 19.05.2020 меняем их местами обратно как было изначально
                                            //rqd.accountableSeries = rqd.userSeries;
                                            //rqd.accountableNumber = rqd.userNumber;
                                            ExecuteScalar($"select username from aspnet_users where id=(select user_id from storagedocs where id=(select id_act from storagedocs where id={id_doc}))", ref obj, null);
                                            string passport = sc.UserPassport((string)obj);
                                            //rqd.userSeries = sc.PassportSeries(passport);
                                            //rqd.userNumber = sc.PassportNumber(passport);
                                            rqd.accountableSeries = sc.PassportSeries(passport);
                                            rqd.accountableNumber = sc.PassportNumber(passport);
                                        }

                                        foreach (BaseProductType value in Enum.GetValues(typeof(BaseProductType)))
                                        {
                                            int count = 0;
                                            if (value == BaseProductType.None) continue;
                                            if (value == BaseProductType.MasterCard) count = bs.countMasterCard[indx];
                                            if (value == BaseProductType.VisaCard) count = bs.countVisaCard[indx];
                                            if (value == BaseProductType.ServiceCard) count = bs.countServiceCard[indx];
                                            if (value == BaseProductType.NFCCard) count = bs.countNFCCard[indx];
                                            if (value == BaseProductType.MirCard) count = bs.countMirCard[indx];
                                            if (value == BaseProductType.PinConvert) count = bs.countPinConvert[indx];
                                            if (count > 0)
                                            {
                                                rqd.materialValue = BranchStore.baseProductTypeToGuid(value);
                                                rqd.count = count;
                                                RMKResponseData rsd = r.RMKCardPersoData(rqd);
                                                if (rsd.status != true)
                                                {
                                                    //13.04.20 в случае ошибок рмк для какой-то проводки, удаляем получившиеся проводки по этому документу
                                                    WebLog.LogClass.WriteToLog("Откат проводок RMK id_doc=" + id_doc.ToString());
                                                    List<RMKData> rmkdata = BranchStore.getRMKData(id_doc, conn);
                                                    RMKRequestData rollback = new RMKRequestData();
                                                    for (int i = 0; i < rmkdata.Count; i++)
                                                    {
                                                        rollback.operationId = rmkdata[i].guidrmk;
                                                        r.RMKCardPersoData(rollback);
                                                    }
                                                    BranchStore.clearRMKData(id_doc, conn);
                                                    throw new Exception(rsd.message);
                                                }
                                                BranchStore.addRMKData(id_doc, (int)value, rsd.operationId, conn);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        WebLog.LogClass.WriteToLog("Удаление RMK Data id_doc=" + id_doc.ToString());
                                        List<RMKData> rmkdata = BranchStore.getRMKData(id_doc, conn);
                                        RMKRequestData rqd = new RMKRequestData();
                                        for (int i = 0; i < rmkdata.Count; i++)
                                        {
                                            rqd.operationId = rmkdata[i].guidrmk;
                                            r.RMKCardPersoData(rqd);
                                        }

                                        BranchStore.clearRMKData(id_doc, conn);
                                    }
                                }
                            }
                        }
                        catch (Exception erm)
                        {
                            string ermes = "Ошибка при обращении к РМК сервису: " + erm.Message;
                            WebLog.LogClass.WriteToLog(ermes);
                            Refr(0, false);
                            ClientScript.RegisterClientScriptBlock(GetType(), "errRMK", "<script type='text/javascript'>$(document).ready(function(){ ShowMessage('" + ermes + "');});</script>");
                            lbInform.Text = erm.Message;
                            return;
                        }

                        #endregion

                        #region кладем файл курьерской службы для персонализации и отправки в филиал
                        try
                        {
                            //поскольку это делается не для всех филиалов, то проверяем еще и на кол-во карт в документе
                            if ((id_type == 8 || id_type == 5) && documentJson.cards.Count > 0) 
                            {
                                documentJson.documentId = number_doc.ToString();
                                if (sost == 0)
                                {
                                    documentJson.operationType = "Confirm";
                                    documentJson.cardsState = (id_type == 8) ? "Хранилище" : "Курьер филиал";
                                }
                                if (sost == 1)
                                {
                                    documentJson.operationType = "Reject";
                                    documentJson.cardsState = (id_type == 8) ? "Обработка" : "Хранилище";
                                }
                                DateTime dtjson = DateTime.Now;
                                //DateTimeOffset dtjsonoff = new DateTimeOffset(dtjson, TimeZoneInfo.Local.GetUtcOffset(dtjson));
                                documentJson.dateTime = dtjson.ToString("yyyy-MM-ddTHH:mmK");
                                var serializer = new Newtonsoft.Json.JsonSerializer();
                                var stringWriter = new StringWriter();
                                using (var writer = new Newtonsoft.Json.JsonTextWriter(stringWriter))
                                {
                                    writer.QuoteName = false;
                                    writer.Formatting = Newtonsoft.Json.Formatting.Indented;
                                    serializer.Serialize(writer, documentJson);
                                }
                                var filename = Path.Combine(ConfigurationManager.AppSettings["CourierFolder"], $"{number_doc}_{dtjson.ToString("yyyyMMddHHmm")}.json");
                                using (StreamWriter sw = new StreamWriter(filename))
                                {
                                    sw.WriteLine(stringWriter.ToString());
                                    sw.Close();
                                }
                            }
                        }
                        catch (Exception exp)
                        {
                            WebLog.LogClass.WriteToLog(exp.ToString());
                            lbInform.Text = "Ошибка создания реестра курьерской службы";
                            return;
                        }
                        #endregion

                        if (sost == 0 && id_type == (int)TypeDoc.PersoCard)
                            SendConfirm(Restrictions.ConfirmPerso, -1);
                        UpdateStorage(id_doc, id_type, sost, dt, null);
                        ChangeGen(id_doc, gen, dt, null);
                        if (sost == 0 && id_type == (int)TypeDoc.PersoCard)
                            SendConfirm(Restrictions.InformProduction, id_doc); // посылаем после подтверждения, чтобы уже была указана дата производства
                        WebLog.LogClass.WriteToLog(sc.BranchId(User.Identity.Name) + " " + User.Identity.Name + " " + String.Format("Документ подтверждение {0} {1} {2}", sc.UserGuid(User.Identity.Name), id_doc, sost));

                        try
                        {
                            if (id_type == (int)TypeDoc.ReceiveToFilial && !isForKilling)
                            {
                                string[] excludesString = System.Configuration.ConfigurationManager.AppSettings["SmsExcludeBranch"].Split(',');
                                if (!excludesString.Contains(SingleQuery.BranchName(current_branch_id)))
                                {
                                    #region сервис информирования по sms ozeki

                                    try
                                    {
                                        Ozeki O = new Ozeki(conn);
                                        if (gen != 0)
                                        {
                                            SmsInfoSection smss =
                                                (SmsInfoSection) ConfigurationManager.GetSection("sms_info");
                                            O.deleteData(id_doc);
                                            Ozekidata[] odata = null;
                                            odata = O.newData(id_doc);
                                            List<Ozekidata> dataforsend = null;
                                            if (odata.Length > 0)
                                            {
                                                //получили все карты
                                                List<Ozekidata> ldata = odata.ToList<Ozekidata>();
                                                //05.06.2020 - раньше удаляли из массива все карты не того продукта, не физиков и без телефона вот таким методом. Теперь в цикле проверяем на настройки из web.config
                                                ////ldata.RemoveAll(m =>
                                                ////    (((m.bin.Trim() == "5209858" && m.prefix_ow.Trim() == "GEN") ||
                                                ////        (m.bin.Trim() == "4000793" && m.prefix_ow.Trim() == "UDK") ||
                                                ////        (m.bin.Trim() == "4000793" && m.prefix_ow.Trim() == "CRD") ||
                                                ////        (m.bin.Trim() == "5209858" && m.prefix_ow.Trim() == "CRD") ||
                                                ////        (m.bin.Trim() == "5209858" && m.prefix_ow.Trim() == "AUR")
                                                ////        ) &&
                                                ////        m.msgout.receiver != null &&
                                                ////        m.company.Trim() == "F" &&
                                                ////        m.id_branchcard == current_branch_id) == false);

                                                // замена текста сообщения
                                                //<Имя-Отчество>, polychite karty "Generation" v otdelenii Banka g. Kazan, ul.Modelnaya, d.10 
                                                //<Имя-Отчество>, polychite karty "Evolution" v otdelenii Banka g. Kazan, ul.Modelnaya, d.10
                                                
                                                dataforsend = new List<Ozekidata>();
                                                for (int i = 0; i < ldata.Count; i++)
                                                {
                                                    //05.06.2020 -  ищем подходящий продукт в web.config
                                                    //сразу откидываем те, которые без телефона и которые еще не в своем конечном филиале
                                                    if (ldata[i].msgout.receiver == null || ldata[i].id_branchcard != current_branch_id)
                                                        continue;
                                                    SmsInfoElement currentSms = null;
                                                    foreach (SmsInfoElement sms in smss.SmsInfo)
                                                    {
                                                        if (ldata[i].bin == sms.Bin && ldata[i].prefix_ow == sms.Code)
                                                        {
                                                            currentSms = sms;
                                                            break;
                                                        }
                                                    }
                                                    // если не нашли подходящий продукт, то ничего не делаем
                                                    //WebLog.LogClass.WriteToLog("CompanyField: '{0}', company: '{1}'", smss.SmsInfo.CompanyField, ldata[i].company);
                                                    if (currentSms == null)
                                                        continue;
                                                    //проверяем, если у этой карты стоит признак, что смс только для физиков, то правильное ли поле Company
                                                    if (currentSms.AllCards.ToLower() == "false" && !smss.SmsInfo.CompanyField.Split(',').Contains(ldata[i].company))
                                                        continue;

                                                    string adress = ldata[i].adress;
                                                    if (adress == null || adress.Trim().Length < 1) adress = " ---";
                                                    else
                                                    {
                                                        string[] a = adress.Split(',');
                                                        if (a.Length > 1)
                                                        {
                                                            long n = 0;
                                                            adress = "";
                                                            if (Int64.TryParse(a[0].Trim(), out n) == true)
                                                            {
                                                                for (int j = 1; j < a.Length; j++)
                                                                {
                                                                    adress += a[j];
                                                                    adress += " ";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                for (int j = 0; j < a.Length; j++)
                                                                {
                                                                    adress += a[j];
                                                                    adress += " ";
                                                                }
                                                            }
                                                            adress = adress.Trim();
                                                        }
                                                    }

                                                    string fio = "";
                                                    //05.06.2020 убрал формирование текста строкой ниже
                                                    //fio = string.Format(", {0} {1}", Ozeki_CardName(ldata[i].prefix_ow.Trim()), adress);
                                                    string[] fiom = ldata[i].fio.Trim().Split(' ');
                                                    if (fiom.Length > 2)
                                                        fio = fiom[1].Trim() + " " + fiom[2].Trim() + fio;
                                                    else fio = ldata[i].fio.Trim() + fio;

                                                    //ldata[i].msgout.msg = fio;
                                                    ldata[i].msgout.msg =
                                                        $"{fio}, получите Вашу карту {currentSms.ShortName} *{ldata[i].last4} в офисе Банка {adress}";
                                                    dataforsend.Add(ldata[i]);
                                                }
                                                odata = ldata.ToArray();
                                            }
                                            //05.06.2020 odata заменено на dataforsend
                                            if (dataforsend != null && dataforsend.Count > 0)
                                            {
                                                O.saveData(dataforsend.ToArray());
                                                WebLog.LogClass.WriteToLog(
                                                    "Для документа id = {0} сформированы данные для OZEKI, кол-во: {1}",
                                                    id_doc, dataforsend.Count);
                                            }
                                            else
                                            {
                                                WebLog.LogClass.WriteToLog(
                                                    "Для документа id = {0} нет данных для OZEKI", id_doc);
                                            }
                                        }
                                        else
                                        {
                                            O.deleteData(id_doc);
                                            WebLog.LogClass.WriteToLog("Удалены OZEKI для документа id = {0}",
                                                id_doc);
                                        }
                                    }
                                    catch (Exception erms)
                                    {
                                        WebLog.LogClass.WriteToLog(
                                            "Ошибка операции с  OZEKI для документа id = {0}: {1}", id_doc, erms.Message);
                                    }

                                    #endregion
                                }
                            }
                        }
                        catch (Exception erm)
                        {
                            string ermes = "Ошибка при обращении к OZEKI: " + erm.Message;
                            WebLog.LogClass.WriteToLog(ermes);
                            Refr(0, false);
                            ClientScript.RegisterClientScriptBlock(GetType(), "errOZEKI", "<script type='text/javascript'>$(document).ready(function(){ ShowMessage('" + ermes + "');});</script>");
                            lbInform.Text = erm.Message;
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        WebLog.LogClass.WriteToLog(ex.ToString());
                    }
                }
                else
                {
                    DataSet ds1 = new DataSet();
                    ds1.Clear();
                    res = ExecuteQuery(String.Format("select id,type,priz_gen,number_doc from StorageDocs where id_deliv={0}", id_deliv.ToString()), ref ds1, null);
                    
                    for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                    {
                        id_doc = Convert.ToInt32(ds1.Tables[0].Rows[i]["id"]);
                        id_type = Convert.ToInt32(ds1.Tables[0].Rows[i]["type"]);
                        sost = Convert.ToInt32(ds1.Tables[0].Rows[i]["priz_gen"]);
                        if (Convert.ToInt32(ds1.Tables[0].Rows[i]["priz_gen"]) == 0)
                            gen = 1;
                    documentJson.Clear();
                    if (!CheckProducts(id_doc, id_type, sost)) return;
                    #region кладем файл курьерской службы для персонализации и отправки в филиал
                    try
                    {
                        number_doc = Convert.ToInt32(ds1.Tables[0].Rows[i]["number_doc"]);
                        if ((id_type == 8 || id_type == 5) && documentJson.cards.Count > 0)
                        {
                            documentJson.documentId = number_doc.ToString();
                            if (sost == 0)
                            {
                                documentJson.operationType = "Confirm";
                                documentJson.cardsState = (id_type == 8) ? "Хранилище" : "Курьер филиал";
                            }
                            if (sost == 1)
                            {
                                documentJson.operationType = "Reject";
                                documentJson.cardsState = (id_type == 8) ? "Обработка" : "Хранилище";
                            }
                            DateTime dtjson = DateTime.Now;
                            //DateTimeOffset dtjsonoff = new DateTimeOffset(dtjson, TimeZoneInfo.Local.GetUtcOffset(dtjson));
                            documentJson.dateTime = dtjson.ToString("yyyy-MM-ddTHH:mmK");
                            var serializer = new Newtonsoft.Json.JsonSerializer();
                            var stringWriter = new StringWriter();
                            using (var writer = new Newtonsoft.Json.JsonTextWriter(stringWriter))
                            {
                                writer.QuoteName = false;
                                writer.Formatting = Newtonsoft.Json.Formatting.Indented;
                                serializer.Serialize(writer, documentJson);
                            }
                            var filename = Path.Combine(ConfigurationManager.AppSettings["CourierFolder"], $"{number_doc}_{dtjson.ToString("yyyyMMddHHmm")}.json");
                            using (StreamWriter sw = new StreamWriter(filename))
                            {
                                sw.WriteLine(stringWriter.ToString());
                                sw.Close();
                            }
                        }
                    }
                    catch (Exception exp)
                    {
                        WebLog.LogClass.WriteToLog(exp.ToString());
                        lbInform.Text = "Ошибка создания реестра курьерской службы";
                        return;
                    }
                    #endregion
                    }

                    SqlTransaction trans = conn.BeginTransaction();
                    try
                    {


                        for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                        {
                            id_doc = Convert.ToInt32(ds1.Tables[0].Rows[i]["id"]);
                            id_type = Convert.ToInt32(ds1.Tables[0].Rows[i]["type"]);
                            sost = Convert.ToInt32(ds1.Tables[0].Rows[i]["priz_gen"]);
                            if (Convert.ToInt32(ds1.Tables[0].Rows[i]["priz_gen"]) == 0)
                                gen = 1;
                            UpdateStorage(id_doc, id_type, sost, dt, trans);
                            ChangeGen(id_doc, gen, dt, trans);
                        }
                        trans.Commit();
                        lbInform.Text = "";
                    }
                    catch (Exception ex)
                    {
                        lbInform.Text = $"Ошибка подтверждения: {ex.Message}";
                        trans.Rollback();
                        WebLog.LogClass.WriteToLog("StorDoc.bSostD_Click ERROR id_doc={0}, id_type={1}, sost = {2}: {3} ", id_doc, id_type, sost, ex.Message);
                        return;
                    }
                }
                WebLog.LogClass.WriteToLog("StorDoc.bSostD_Click End id_doc={0}, id_type={1}, sost = {2}, user = {3}, userbranch = {4}", id_doc, id_type, sost, User.Identity.Name, sc.BranchId(User.Identity.Name));
                Refr(0, false);
            }
        }
        private string Ozeki_CardName(string owprefix)
        {
            switch (owprefix)
            {
                case ("GEN"):
                    return "получите карту Generation в отделении Банка";
                case ("UDK"):
                    return "получите карту Evolution в отделении Банка";
                case ("CRD"):
                    return "кредитная карта Emotion готова к выдаче. Получите в офисе Банка";
                case ("AUR"):
                    return "получите карту Aurum в отделении Банка";
                default:
                    return "получите карту в отделении Банка";
            }
        }
        private int GetPerson(int id_act)
        {
            DataSet ds = new DataSet();
            ds.Clear();
            res = ExecuteQuery(String.Format("select top 1 * from AccountablePerson_StorageDocs where id_doc={0} order by id desc", id_act), ref ds, null);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                return Convert.ToInt32(ds.Tables[0].Rows[0]["id_person"]);
            }
            else return 0;
        }
        private void ExcelReestrFilial(int id_sd, ExcelAp ep)
        {
            ExcelReestrFilial(id_sd, ep, null, null);
        }

        private void ExcelReestrFilial(int id_sd, ExcelAp ep, BranchStore[] excludeBranchs, BranchStore onlyBranch)
        {
            string dt_doc = "";

            dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now, "г.");
            ep.SetText(6, 5, dt_doc);

            DataSet ds1 = new DataSet();
            ds1.Clear();
            //res = ExecuteQuery(String.Format("select fio,pan,company,IdentDepInit,IdentDepCard from V_Cards inner join Cards_StorageDocs on V_Cards.id = Cards_StorageDocs.id_card where id_doc={0} order by IdentDepCard, company, fio", id_sd.ToString()), ref ds1, null);
            res = ExecuteQuery(String.Format("SELECT TOP (100) PERCENT dbo.Cards.id_branchCard, dbo.Cards.fio, dbo.Cards.pan, dbo.Cards.company, dbo.Branchs.ident_dep AS identDepInit, Branchs_1.ident_dep AS identDepCard FROM dbo.Cards INNER JOIN dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card INNER JOIN dbo.Branchs ON dbo.Cards.id_branchInit = dbo.Branchs.id INNER JOIN dbo.Branchs AS Branchs_1 ON dbo.Cards.id_branchCard = Branchs_1.id WHERE (dbo.Cards_StorageDocs.id_doc = {0}) ORDER BY dbo.Cards.company, dbo.Cards.fio", id_sd.ToString()), ref ds1, null);
            ep.SetFormat(9, 1, 9 + ds1.Tables[0].Rows.Count, 6, "@");
            int skiprows = 0;
            for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
            {
                if (excludeBranchs != null && excludeBranchs.Length > 0)
                {
                    BranchStore b = excludeBranchs.FirstOrDefault(p => p.id == Convert.ToInt32(ds1.Tables[0].Rows[i]["id_branchCard"]));
                    if (b != null)
                    {
                        skiprows++; continue;
                    }
                }
                if (onlyBranch != null)
                {
                    if (onlyBranch.id != Convert.ToInt32(ds1.Tables[0].Rows[i]["id_branchCard"]))
                    {
                        skiprows++; continue;
                    }
                }
                ep.SetText(9 + i - skiprows, 1, (i + 1 - skiprows).ToString());
                ep.SetText(9 + i - skiprows, 2, ds1.Tables[0].Rows[i]["fio"].ToString().Trim());
                ep.SetText(9 + i - skiprows, 3, ds1.Tables[0].Rows[i]["pan"].ToString().Trim());
                ep.SetText(9 + i - skiprows, 4, ds1.Tables[0].Rows[i]["company"].ToString().Trim());
                ep.SetText(9 + i - skiprows, 5, ds1.Tables[0].Rows[i]["IdentDepInit"].ToString().Trim());
                ep.SetText(9 + i - skiprows, 6, ds1.Tables[0].Rows[i]["IdentDepCard"].ToString().Trim());
            }
            ep.SetRangeBold(8, 1, 8, 6);
            //            ep.SetRangeAutoFit(8, 1, 8 + ds1.Tables[0].Rows.Count, 6); // чтобы влазило на страницу
            ep.SetRangeBorders(8, 1, 8 + ds1.Tables[0].Rows.Count - skiprows, 6);
        }

        
        private void Excel2AktReestrPerso(int doc_id, int br_id, string branch, int page, ArrayList prods, ExcelAp ep)
        {
            string dt_doc = "";
            dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
            // страницы реестра
            ep.SetWorkSheet(page + 3);
            ep.SetText(6, 5, dt_doc);
            DataSet ds1 = new DataSet();
            ds1.Clear();
            //res = ExecuteQuery(String.Format("select id_prb, fio,pan,company,IdentDepInit,IdentDepCard from V_Cards inner join Cards_StorageDocs on V_Cards.id = Cards_StorageDocs.id_card where id_doc={0} and id_branchcard={1} order by company, fio", doc_id, br_id), ref ds1, null);
            res = ExecuteQuery(String.Format("SELECT dbo.Cards.id_prb, dbo.Cards.isPin, dbo.Cards.fio, dbo.Cards.pan, dbo.Cards.company, dbo.Branchs.ident_dep AS identDepInit, Branchs_1.ident_dep AS identDepCard FROM dbo.Cards INNER JOIN dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card INNER JOIN dbo.Branchs ON dbo.Cards.id_branchInit = dbo.Branchs.id INNER JOIN dbo.Branchs AS Branchs_1 ON dbo.Cards.id_branchCard = Branchs_1.id WHERE (dbo.Cards_StorageDocs.id_doc = {0}) AND (dbo.Cards.id_branchCard = {1}) ORDER BY dbo.Cards.company, dbo.Cards.fio", doc_id, br_id), ref ds1, null);
            ep.SetFormat(9, 1, 9 + ds1.Tables[0].Rows.Count, 6, "@");
            for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
            {
                ep.SetText(9 + i, 1, (i + 1).ToString());
                ep.SetText(9 + i, 2, ds1.Tables[0].Rows[i]["fio"].ToString().Trim());
                ep.SetText(9 + i, 3, ds1.Tables[0].Rows[i]["pan"].ToString().Trim());
                ep.SetText(9 + i, 4, ds1.Tables[0].Rows[i]["company"].ToString().Trim());
                ep.SetText(9 + i, 5, ds1.Tables[0].Rows[i]["IdentDepInit"].ToString().Trim());
                ep.SetText(9 + i, 6, ds1.Tables[0].Rows[i]["IdentDepCard"].ToString().Trim());
                foreach (MyProd mp in prods)
                    if (mp.ID == Convert.ToInt32(ds1.Tables[0].Rows[i]["id_prb"]))
                    {
                        mp.cnts[0] += 1;
                        if(Convert.ToInt32(ds1.Tables[0].Rows[i]["isPin"])>0)
                            mp.cnts[1] += 1;
                    }
            }
            ep.SetRangeBold(8, 1, 8, 6);
            ep.SetRangeBorders(8, 1, 8 + ds1.Tables[0].Rows.Count, 6);
            // акт страница 1
            string fio = GetFilialPeople(br_id, null);
            for (int tt = 0; tt < 2; tt++)
            {
                ep.SetWorkSheet(page + tt + 1);
                ep.SetText(7, 1, "Дата: " + dt_doc);
                ep.SetText(71, 1, dt_doc);
                ep.SetText_Name("fio", fio);
                ep.SetText(4, 3, branch);
                ep.SetText(66, 6, branch);
                int j = 0, cnt = 0, cntpin = 0;
                foreach (MyProd mp in prods)
                {
                    if (mp.Type == 1 && mp.cnts[0] > 0)
                    {
                        ep.SetText((++j) + 10, 1, "  " + j.ToString() + ".   " + mp.Name);
                        ep.SetText(j + 10, 6, mp.cnts[0].ToString());
                        cnt += mp.cnts[0];
                        cntpin += mp.cnts[1];
                    }
                }
                ep.SetText((++j) + 10, 1, "   Всего");
                ep.SetText(j + 10, 6, cnt.ToString());
                ep.ShowRows(10, j + 10 + 5);
                ep.SetText(58, 6, cntpin.ToString()); //пин конверты
                j = 0; cnt = 0;
                foreach (MyProd mp in prods)
                {
                    if (mp.Type == 3 && mp.cnts[0] > 0)
                    {
                        ep.SetText((++j) + 25, 1, "  " + j.ToString() + ".  " + mp.Name);
                        ep.SetText(j + 50, 6, mp.cnts[0].ToString());
                        cnt += mp.cnts[0];
                    }
                }
                ep.SetText((++j) + 50, 1, "   Всего");
                ep.SetText(j + 50, 6, cnt.ToString());
            }
        }

        private void ExcelDelivFilial(int id_sd, int id_branch, string branch, ExcelAp ep)
        {
            ExcelDelivFilial(id_sd, id_branch, branch, (int)TypeDoc.SendToFilial, ep, null, null);
        }

        private void ExcelDelivFilial(int id_sd, int id_branch, string branch, ExcelAp ep, BranchStore[] excludeBranchs, BranchStore onlyBranch)
        {
            ExcelDelivFilial(id_sd, id_branch, branch, (int)TypeDoc.SendToFilial, ep, excludeBranchs, onlyBranch);
        }

        private void ExcelDelivFilial(int id_sd, int id_branch, string branch, int id_type, ExcelAp ep)
        {
            ExcelDelivFilial(id_sd, id_branch, branch, id_type, ep, null, null);
        }

        private void ExcelDelivFilial(int id_sd, int id_branch, string branch, int id_type, ExcelAp ep, BranchStore[] excludeBranchs, BranchStore onlyBranch)
        {
            string dt_doc = "";
            int all_cnt_n = 0; int all_cnt_p = 0; int all_cnt_pin = 0;
            int cnt_prod = 0; int cnt_dop_prod = 0;
            int cnt = 0;
            int type_p = 0;
            string prod_name = "";

            dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
            ep.SetText(7, 1, "Дата: " + dt_doc);


            //            string fio = GetFilialPeople(id_branch, null);
            ep.SetText_Name("fio", GetFilialPeople(id_branch, null));


            string position = GetPosition();
            if (position != null) ep.SetText_Name("PERSON_FROM", position);
            string positionFio = GetPositionFio();
            if (positionFio != null) ep.SetText_Name("PERSON_FROM_FIO", "____________/ " + positionFio + " /");

            if (id_type == (int)TypeDoc.SendToFilialFilial)
            {
                ep.SetText(4, 1, branch);
                //ep.SetText(46, 1, dt_doc);

                if (branch.IndexOf("-->") > 0) // то это пересылка между филиалами
                {
                    ep.SetText(4, 1, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[1]);
                }
            }
            else
            {
                ep.SetText(4, 3, branch);
                ep.SetText(66, 6, branch);
                ep.SetText(71, 1, dt_doc);


                if (branch.IndexOf("-->") > 0) // то это пересылка между филиалами
                {
                    ep.SetText(4, 3, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[1]);
                    ep.SetText(66, 6, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[1]);
                    ep.SetText(66, 1, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[0]);
                }
            }
            DataSet ds1 = new DataSet();

            ds1.Clear();
            res = ExecuteQuery(String.Format("select prod_name,cnt_new,cnt_perso,id_type,id_prb from V_Products_StorageDocs where id_doc={0} order by id_type,id_sort,prod_name", id_sd.ToString()), ref ds1, null);

            int skipprod = 0;
            for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
            {
                type_p = Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]);
                int id_prb_c = Convert.ToInt32(ds1.Tables[0].Rows[j]["id_prb"]);
                if (type_p == 1)
                {
                    prod_name = ds1.Tables[0].Rows[j]["prod_name"].ToString();
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_perso"]);

                    if (excludeBranchs != null && excludeBranchs.Length > 0)
                    {
                        DataSet dsex = new DataSet();
                        for (int b = 0; b < excludeBranchs.Length; b++)
                        {
                            dsex.Clear();
                            object obj = null;
                            ExecuteScalar(String.Format("select count(c.id_BranchCard) as cntminus " +
                                                        " from Cards_StorageDocs cs left join Cards c on cs.id_card=c.id " +
                                                        " where id_doc={0} and id_BranchCard={1} and id_prb={2}",
                                                        id_sd.ToString(), excludeBranchs[b].id, Convert.ToInt32(ds1.Tables[0].Rows[j]["id_prb"])), ref obj, null);
                            if (obj != DBNull.Value) cnt -= Convert.ToInt32(obj);
                            obj = null;
                            ExecuteScalar(String.Format("select count(c.id_BranchCard) as cntminus " +
                                                        " from Cards_StorageDocs cs left join Cards c on cs.id_card=c.id " +
                                                        " where id_doc={0} and id_BranchCard={1} and id_prb={2} and c.ispin>0",
                                                        id_sd.ToString(), excludeBranchs[b].id, Convert.ToInt32(ds1.Tables[0].Rows[j]["id_prb"])), ref obj, null);
                            if (obj != DBNull.Value) all_cnt_pin -= Convert.ToInt32(obj);
                        }
                        if (cnt <= 0)
                        {
                            skipprod++; continue;
                        }
                    }
                    if (onlyBranch != null)
                    {

                        DataSet dsex = new DataSet();
                        dsex.Clear();
                        object obj = null;
                        ExecuteScalar(String.Format("select count(c.id_BranchCard) " +
                                                    " from Cards_StorageDocs cs left join Cards c on cs.id_card=c.id " +
                                                    " where id_doc={0} and id_BranchCard!={1} and id_prb={2}",
                                                    id_sd.ToString(), onlyBranch.id, Convert.ToInt32(ds1.Tables[0].Rows[j]["id_prb"])), ref obj, null);
                        if (obj != DBNull.Value) cnt -= Convert.ToInt32(obj);
                        obj = null;
                        ExecuteScalar(String.Format("select count(c.id_BranchCard) " +
                                                    " from Cards_StorageDocs cs left join Cards c on cs.id_card=c.id " +
                                                    " where id_doc={0} and id_BranchCard!={1} and id_prb={2} and c.ispin>0",
                                                    id_sd.ToString(), onlyBranch.id, Convert.ToInt32(ds1.Tables[0].Rows[j]["id_prb"])), ref obj, null);
                        if (obj != DBNull.Value) all_cnt_pin -= Convert.ToInt32(obj);
                        if (cnt <= 0)
                        {
                            skipprod++; continue;
                        }
                    }

                    ep.SetText(cnt_prod + 10, 1, "  " + (j + 1 - skipprod).ToString() + ".   " + prod_name);
                    if (id_type == (int)TypeDoc.SendToFilialFilial)
                        ep.SetText(cnt_prod + 10, 2, cnt.ToString());
                    else
                        ep.SetText(cnt_prod + 10, 6, cnt.ToString());

                    all_cnt_p = all_cnt_p + cnt;
                    cnt_prod = cnt_prod + 1;
                }
                if (type_p == 2)
                {
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_perso"]);
                    all_cnt_pin = all_cnt_pin + cnt;
                }
                if (type_p == 3 && id_type != (int)TypeDoc.SendToFilialFilial)
                {
                    prod_name = ds1.Tables[0].Rows[j]["prod_name"].ToString();
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_new"]);
                    if (excludeBranchs != null && excludeBranchs.Length > 0)
                    {
                        String exBranch = "";
                        for (int b = 0; b < excludeBranchs.Length; b++)
                        {
                            if (exBranch.Length > 0) exBranch += ",";
                            exBranch += excludeBranchs[b].id.ToString();
                        }
                        DataSet dsex = new DataSet();
                        for (int m = 0; m < ds1.Tables[0].Rows.Count; m++)
                        {
                            if (1 != Convert.ToInt32(ds1.Tables[0].Rows[m]["id_type"])) continue;
                            int id_prb_p = Convert.ToInt32(ds1.Tables[0].Rows[m]["id_prb"]);
                            dsex.Clear();
                            object obj = null;
                            ExecuteScalar(String.Format("select sum(pa.cnt) " +
                                                        " from Cards_StorageDocs cs" +
                                                        " left join Cards c on cs.id_card=c.id" +
                                                        " join Products_Attachments pa on pa.id_prb_parent=c.id_prb" +
                                                        " where cs.id_doc={0}" +
                                                        " and c.id_BranchCard in({1})" +
                                                        " and pa.id_prb={2}" +
                                                        " and pa.id_prb_parent={3}", id_sd, exBranch, id_prb_c, id_prb_p), ref obj, null);
                            if (obj != DBNull.Value) 
                                cnt -= Convert.ToInt32(obj);
                        }
                        if (cnt < 0) cnt = 0; // такого быть не должно
                    }
                    if (onlyBranch != null)
                    {
                        cnt = 0;
                        DataSet dsex = new DataSet();
                        for (int m = 0; m < ds1.Tables[0].Rows.Count; m++)
                        {
                            if (1 != Convert.ToInt32(ds1.Tables[0].Rows[m]["id_type"])) continue;
                            int id_prb_p = Convert.ToInt32(ds1.Tables[0].Rows[m]["id_prb"]);
                            dsex.Clear();
                            object obj = null;
                            ExecuteScalar(String.Format("select sum(pa.cnt) " +
                                                        " from Cards_StorageDocs cs" +
                                                        " left join Cards c on cs.id_card=c.id" +
                                                        " join Products_Attachments pa on pa.id_prb_parent=c.id_prb" +
                                                        " where cs.id_doc={0}" +
                                                        " and c.id_BranchCard={1}" +
                                                        " and pa.id_prb={2}" +
                                                        " and pa.id_prb_parent={3}", id_sd, onlyBranch.id, id_prb_c, id_prb_p), ref obj, null);
                            if (obj != DBNull.Value)
                                cnt += Convert.ToInt32(obj);
                        }
                    }

                    ep.SetText(cnt_dop_prod + 50, 1, "  " + (cnt_dop_prod + 1).ToString() + ".   " + prod_name);
                    ep.SetText(cnt_dop_prod + 50, 6, cnt.ToString());

                    cnt_dop_prod = cnt_dop_prod + 1;
                    all_cnt_n = all_cnt_n + cnt;
                }
            }

            ep.SetText(cnt_prod + 10, 1, "   Всего");
            if (id_type == (int)TypeDoc.SendToFilialFilial)
            {
                ep.SetText(cnt_prod + 10, 2, all_cnt_p.ToString());
                ep.SetText(40, 2, all_cnt_pin.ToString());
                ep.ShowRows(10, 10 + cnt_prod + 2);
            }
            else
            {
                ep.SetText(cnt_prod + 10, 6, all_cnt_p.ToString());
                ep.SetText(cnt_dop_prod + 50, 1, "   Всего");
                ep.SetText(cnt_dop_prod + 50, 6, all_cnt_n.ToString());
                ep.SetText(58, 6, all_cnt_pin.ToString());
                ep.ShowRows(10, 10 + cnt_prod + 5);
            }
        }

        /*
        private void ExcelDelivFilial(int id_sd, int id_branch, string branch, int id_type, ExcelAp ep)
        {
            string dt_doc = "";
            int all_cnt_n = 0; int all_cnt_p = 0; int all_cnt_pin = 0;
            int cnt_prod = 0; int cnt_dop_prod = 0;
            int cnt = 0;
            int type_p = 0;
            string prod_name = "";

            dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
            ep.SetText(7, 1, "Дата: " + dt_doc);


            //            string fio = GetFilialPeople(id_branch, null);
            ep.SetText_Name("fio", GetFilialPeople(id_branch, null));


            string position = GetPosition();
            if (position != null) ep.SetText_Name("PERSON_FROM", position);
            string positionFio = GetPositionFio();
            if (positionFio != null) ep.SetText_Name("PERSON_FROM_FIO", "____________/ " + positionFio + " /");

            if (id_type == (int)TypeDoc.SendToFilialFilial)
            {
                ep.SetText(4, 1, branch);
                //ep.SetText(46, 1, dt_doc);

                if (branch.IndexOf("-->") > 0) // то это пересылка между филиалами
                {
                    ep.SetText(4, 1, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[1]);
                }
            }
            else
            {
                ep.SetText(4, 3, branch);
                ep.SetText(66, 6, branch);
                ep.SetText(71, 1, dt_doc);


                if (branch.IndexOf("-->") > 0) // то это пересылка между филиалами
                {
                    ep.SetText(4, 3, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[1]);
                    ep.SetText(66, 6, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[1]);
                    ep.SetText(66, 1, branch.Split(new string[] { "-->" }, StringSplitOptions.None)[0]);
                }
            }
            DataSet ds1 = new DataSet();

            ds1.Clear();
            res = ExecuteQuery(String.Format("select prod_name,cnt_new,cnt_perso,id_type from V_Products_StorageDocs where id_doc={0} order by id_type,id_sort,prod_name", id_sd.ToString()), ref ds1, null);
            for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
            {
                type_p = Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]);
                if (type_p == 1)
                {
                    prod_name = ds1.Tables[0].Rows[j]["prod_name"].ToString();
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_perso"]);

                    ep.SetText(cnt_prod + 10, 1, "  " + (j + 1).ToString() + ".   " + prod_name);
                    if (id_type == (int)TypeDoc.SendToFilialFilial)
                        ep.SetText(cnt_prod + 10, 2, cnt.ToString());
                    else
                        ep.SetText(cnt_prod + 10, 6, cnt.ToString());

                    all_cnt_p = all_cnt_p + cnt;
                    cnt_prod = cnt_prod + 1;
                }
                if (type_p == 2)
                {
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_perso"]);
                    all_cnt_pin = all_cnt_pin + cnt;
                }
                if (type_p == 3 && id_type != (int)TypeDoc.SendToFilialFilial)
                {
                    prod_name = ds1.Tables[0].Rows[j]["prod_name"].ToString();
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_new"]);

                    ep.SetText(cnt_dop_prod + 50, 1, "  " + (cnt_dop_prod + 1).ToString() + ".   " + prod_name);
                    ep.SetText(cnt_dop_prod + 50, 6, cnt.ToString());

                    cnt_dop_prod = cnt_dop_prod + 1;
                    all_cnt_n = all_cnt_n + cnt;
                }
            }

            ep.SetText(cnt_prod + 10, 1, "   Всего");
            if (id_type == (int)TypeDoc.SendToFilialFilial)
            {
                ep.SetText(cnt_prod + 10, 2, all_cnt_p.ToString());
                ep.SetText(40, 2, all_cnt_pin.ToString());
                ep.ShowRows(10, 10 + cnt_prod + 2);
            }
            else
            {
                ep.SetText(cnt_prod + 10, 6, all_cnt_p.ToString());
                ep.SetText(cnt_dop_prod + 50, 1, "   Всего");
                ep.SetText(cnt_dop_prod + 50, 6, all_cnt_n.ToString());
                ep.SetText(58, 6, all_cnt_pin.ToString());
                ep.ShowRows(10, 10 + cnt_prod + 5);
            }
        }
        */

        private void ExcelReturnFilial(int id_sd, int id_branch, string branch, ExcelAp ep)
        {
            int id_type = Convert.ToInt32(dListDoc.SelectedItem.Value);

            int b_main_filial = BranchStore.getBranchMainFilial(id_branch,false);
            string s_main_filial="";

            string dt_doc = "";
            int all_cnt_n = 0; int all_cnt_p = 0; int all_cnt_pin = 0;
            int cnt_prod = 0; int cnt_dop_prod = 0;
            int cnt = 0;
            int type_p = 0;
            int cnt_perso = 0, cnt_perso_pin = 0;
            string prod_name = "";

            dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
            ep.SetText(7, 1, "Дата: " + dt_doc);
            ep.SetText(57, 1, dt_doc);

            //            string fio = GetFilialPeople(id_branch, null);
            ep.SetText_Name("fio", GetFilialPeople(id_branch, null));

//            ep.SetText(4, 3, branch);
            if (b_main_filial > 0 && b_main_filial!=id_branch)
            {
                object obj=null;
                ExecuteScalar("select department from Branchs where id=" + b_main_filial.ToString(), ref obj, null);
                if (obj != null) s_main_filial = obj.ToString();
                if  (id_type != (int)TypeDoc.SendToBank) //письмо от 12.12.2018
                    ep.SetText(4, 1, "Представителю (" + s_main_filial + ")");
                ep.SetText(50, 1, "Представитель:");
                ep.SetText(51, 1, sc.UserPosition(User.Identity.Name));
                ep.SetText(52, 1, branch);
                ep.SetText(55, 1, "_____________ / " + sc.UserFIO(User.Identity.Name) + " /");

                ep.SetText(51, 6, "Представитель:");
                if (id_type != (int)TypeDoc.SendToBank) //письмо от 12.12.2018
                    ep.SetText(52, 6, s_main_filial);
                //ep.SetRangeAlignment(52,6,52,10, Microsoft.Office.Interop.Excel.Constants.xlCenter);
                ep.SetRangeBold(52, 6, 52, 6);
                ep.SetText(55, 6, "_____________ / ________________ /");
            }
            else
            {
                if (b_main_filial > 0 && b_main_filial==id_branch)
                {

                    ep.SetText(50, 1, "Представитель:");
                    ep.SetText(51, 1, sc.UserPosition(User.Identity.Name));
                    ep.SetText(52, 1, branch);
                    ep.SetText(55, 1, "_____________ / " + sc.UserFIO(User.Identity.Name) + " /");

                    ep.SetText(55, 6, "_____________ / ________________ /");
                }
                else ep.SetText(52, 1, branch);
            }
            

            DataSet ds1 = new DataSet();

            ds1.Clear();
            res = ExecuteQuery(String.Format("select prod_name,cnt_new,cnt_brak,cnt_perso,id_type from V_Products_StorageDocs where id_doc={0} order by id_type,id_sort,prod_name", id_sd.ToString()), ref ds1, null);
            for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
            {
                type_p = Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]);
                if (type_p == 1)
                {
                    prod_name = ds1.Tables[0].Rows[j]["prod_name"].ToString();
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_brak"]);

                    ep.SetText(cnt_prod + 11, 1, "  " + (j + 1).ToString() + ".   " + prod_name);
                    ep.SetText(cnt_prod + 11, 6, cnt.ToString());

                    all_cnt_p = all_cnt_p + cnt;
                    
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_perso"]);
                    ep.SetText(cnt_prod + 11, 9, cnt.ToString());
                    cnt_perso += cnt;
//                    cnt_perso_pin += cnt;
                    cnt_prod = cnt_prod + 1;

                }
                if (type_p == 2)
                {
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_brak"]);
                    all_cnt_pin = all_cnt_pin + cnt;
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_perso"]);
                    cnt_perso_pin += cnt;
                }
                if (type_p == 3)
                {
                    prod_name = ds1.Tables[0].Rows[j]["prod_name"].ToString();
                    cnt = Convert.ToInt32(ds1.Tables[0].Rows[j]["cnt_brak"]);

                    ep.SetText(cnt_dop_prod + 36, 1, "  " + (cnt_dop_prod + 1).ToString() + ".   " + prod_name);
                    ep.SetText(cnt_dop_prod + 36, 6, cnt.ToString());

                    cnt_dop_prod = cnt_dop_prod + 1;
                    all_cnt_n = all_cnt_n + cnt;
                }
            }

            ep.SetText(cnt_prod + 11, 1, "   Всего");
            ep.SetText(cnt_prod + 11, 6, all_cnt_p.ToString());
            ep.SetText(cnt_prod + 11, 9, cnt_perso.ToString());
            ep.SetText(cnt_dop_prod + 36, 1, "   Всего");
            ep.SetText(cnt_dop_prod + 36, 6, all_cnt_n.ToString()); 
            ep.SetText(44, 6, all_cnt_pin.ToString());
            ep.SetText(44, 9, cnt_perso_pin.ToString());
            ep.ShowRows(11, 11+cnt_prod + 2);

        }
        private void ChangeGen(int id_doc,int gen,DateTime dt, SqlTransaction trans)
        {
            WebLog.LogClass.WriteToLog("StorDoc.ChangeGen Start id_doc={0}, gen = {1}, user = {2}, userbranch = {3}", id_doc, gen, User.Identity.Name, sc.BranchId(User.Identity.Name));                                               
            SqlCommand sqCom = new SqlCommand();
            if (trans != null)
                sqCom.Transaction = trans;
            if (dt == DateTime.MinValue)
                sqCom.CommandText = "update StorageDocs set priz_gen=@priz_gen,user_id=@user_id where id=@id";
            else
                sqCom.CommandText = "update StorageDocs set priz_gen=@priz_gen,user_id=@user_id,date_doc=@dt where id=@id";
            sqCom.CommandText += "\r\nupdate StorageDocs set date_time=CONVERT(datetime,(CONVERT(varchar(15),date_doc,104) + ' ' + time_doc),104) where time_doc is not null and id=@id";
            sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
            sqCom.Parameters.Add("@user_id", SqlDbType.Int).Value = sc.UserId(User.Identity.Name);
            if (dt != DateTime.MinValue)
                sqCom.Parameters.Add("@dt", SqlDbType.DateTime).Value = dt;
            sqCom.Parameters.Add("@priz_gen", SqlDbType.Bit).Value = gen;
            ExecuteNonQuery(sqCom, null);
            //sqCom.Parameters.Clear();
            //sqCom.CommandText = "update StorageDocs set date_time=CONVERT(datetime,(CONVERT(varchar(15),date_doc,104) + ' ' + time_doc),104)s where time_doc is not null";
            WebLog.LogClass.WriteToLog("StorDoc.ChangeGen End id_doc={0}, gen = {1}, user = {2}, userbranch = {3}", id_doc, gen, User.Identity.Name, sc.BranchId(User.Identity.Name));                                               
        }

        protected void bExcelD_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id_type = Convert.ToInt32(dListDoc.SelectedItem.Value);
                if (FuncClass.ClientType == ClientType.AkBars)
                    WorkExcel(id_type);
                if (FuncClass.ClientType == ClientType.Uzcard)
                    WorkExcelUz(id_type);
            }
        }
        private void WorkExcel(int id_type)
        {
            WorkExcel(id_type, true);
        }
        protected void WorkExcel(int id_type, bool send)
        {
            try
            {
                string doc = "Empty.xls";
                int id_deliv = 0;
                userToFilialFilial="";

                if (gvDocs.Rows.Count > 0)
                    id_deliv = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_deliv"]);
                
                if (id_type == (int)TypeFormDoc.AccompDoc) doc = "Attachment40.xls";
                if (id_type == (int)TypeFormDoc.OfficeNote) doc = "Attachment1.xls";
                if (id_type == (int)TypeFormDoc.OfficeNote1) doc = "Attachment1_1.xls";
                if (id_type == (int)TypeDoc.CardFromStorage) doc = "Attachment2.xls";
                if (id_type == (int)TypeDoc.PinFromStorage) doc = "Attachment3.xls";
                if (id_type == (int)TypeDoc.CardToStorage) doc = "Attachment4.xls";
                if (id_type == (int)TypeDoc.PinToStorage) doc = "Attachment5.xls";
                //if ((id_type == (int)TypeDoc.SendToFilial) && (id_deliv==0)) doc = "Attachment6.xls";
                //if (id_type == (int)TypeDoc.SendToFilial) doc = "Attachment6.xls";
                if (id_type == (int)TypeDoc.SendToFilial) doc = "Attachment6_8.xls";
                
                if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToFilialPacket)
                {
                    doc = "Attachment6_8.xls";
                    int id_actsel=Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_act"]);
                    if (id_actsel > 0)
                    {
                        SqlCommand sel = new SqlCommand();
                        sel.CommandText = "select type from StorageDocs where id=" + id_actsel.ToString();
                        object obj = null;
                        ExecuteScalar(sel, ref obj, null);
                        if(obj!=null && Convert.ToInt32(obj)==19)
                        {
                            id_type = (int)TypeDoc.SendToFilialFilial;
                            sel.CommandText = "select LoweredUserName from V_StorageDocs where id=" + id_actsel.ToString();
                            ExecuteScalar(sel, ref obj, null);
                            userToFilialFilial = obj.ToString();
                        }
                    }
                }
                
                if (id_type == (int)TypeDoc.SendToFilialFilial) doc = "Attachment6_f.xls";
                
                if (id_type == (int)TypeDoc.DontReceiveToFilial) doc = "Attachment6_8.xls";
                if (id_type == (int)TypeFormDoc.Register) doc = "Attachment8.xls";
                if (id_type == (int)TypeFormDoc.SortList) doc = "Attachment7.xls";
                if (id_type == (int)TypeDoc.PersoCard) doc = "Attachment2_5.xls";
                if (id_type == (int)TypeFormDoc.PochtaAct) doc = "Attachment11.xls";
                if (id_type == (int)TypeFormDoc.TransportAct) doc = "Attachment12.xls";
                if (id_type == (int)TypeFormDoc.Akt2Reestr) doc = "Attachment6_8.xls";
                if (id_type == (int)TypeFormDoc.Akt2ReestrPerso) doc = "Attachment6_8.xls";
                if (id_type == (int)TypeFormDoc.Courier7777) doc = "Attachment50.xls";
                if (id_type == (int)TypeDoc.DeleteBrak)
                {
                    if(branch_main_filial>0) doc = "Attachment14_f.xls";
                    else doc = "Attachment14.xls";
                }
                if (id_type == (int)TypeDoc.SendToClient) doc = "Attachment16.xls";
                if (id_type == (int)TypeDoc.SendToClientFromPodotchet) doc = "Attachment16.xls";
                if (id_type == (int)TypeDoc.FilialFilial) doc = "Attachment18.xls";
                if (id_type == (int)TypeDoc.SendToBank) doc = "Attachment17.xls";
                if (id_type == (int)TypeDoc.Expendables) doc = "Attachment19.xls";
                if (id_type == (int)TypeDoc.Reklama) doc = "Attachment20.xls";
                if (id_type == (int)TypeDoc.ToWrapping) doc = "Attachment21.xls";
                if (id_type == (int)TypeDoc.FromWrapping) doc = "Attachment22.xls";
                if (id_type == (int)TypeDoc.SendToExpertiza) doc = "Attachment6_8.xls";
                if (id_type == (int)TypeDoc.ReceiveToExpertiza) doc = "Attachment6_8.xls";
                if (id_type == (int)TypeDoc.SendToPodotchet) doc = "Attachment16.xls";
                if (id_type == (int)TypeDoc.ReceiveToPodotchet) doc = "Attachment16.xls";
                if (id_type == (int)TypeDoc.ReceiveFromPodotchet) doc = "Attachment16.xls";
                if (id_type == (int)TypeDoc.ReturnFromPodotchet) doc = "Attachment16.xls";
                if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124 
                    || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124)
                    doc = "Attachment124.xls";
                if (id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz
                    || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz)
                    doc = "Attachment124.xls";
                if (id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                    || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
                    doc = "Attachment16.xls";
                if (id_type == (int) TypeFormDoc.Book124Label)
                    doc = "Attachment124_Label.xls";


                // 01.11.18 для всех документов экспертизы и подотчетных делаем акт+реестр
                //if (id_type >= (int)TypeDoc.ReceiveToFilialExpertiza && id_type <= (int)TypeDoc.WriteOfPodotchet)
                //  doc = "Attachment6_8.xls";

                System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                ExcelAp ep = new ExcelAp();
                if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + doc))
                {
                    //ep.SaveAsDoc(ConfigurationSettings.AppSettings["DocPath"] + "Temp/" + doc, false);
                    if (id_type == (int)TypeDoc.PersoCard)
                    {
                        ep.SetWorkSheet(1);
                        ZapExcel(1, id_type, id_deliv, ep);

                        ep.SetWorkSheet(2);
                        ZapExcel(2, id_type, id_deliv, ep);

                        ep.SetWorkSheet(3);
                        ZapExcel(3, id_type, id_deliv, ep);

                        ep.SetWorkSheet(4);
                        ZapExcel(4, id_type, id_deliv, ep);

                        //ep.SetWorkSheet(5);
                        //ep.SetText(10, 1, String.Format("{0:dd.MM.yyyy}", DateTime.Now));
                        

                    }
                    else
                    {
                        ep.SetWorkSheet(1);
                        ZapExcel(id_type, 0, id_deliv, ep);
                    }

                    if (doc.Length > 0 && WebConfigurationManager.AppSettings["DocPath"] != null)
                    {
                        doc = String.Format("{0}Temp\\{1}", WebConfigurationManager.AppSettings["DocPath"], doc);
                        ep.SaveAsDoc(doc, false);
                        docTempName = doc;
                    }
                }
                ep.Close();
                //ep.Show();
                GC.Collect();
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                if (doc.Length > 0 && send)
                    ep.ReturnXls(Response, doc);
            }
            catch (Exception ex)
            {
                WebLog.LogClass.WriteToLog(ex.ToString());
            }
        }
        protected void WorkExcelUz(int id_type)
        {
            try
            {
                string doc = "Empty.xls";
                if (id_type == (int)TypeDoc.PersoCard) doc = "Uzcard_1.xls";
                if (id_type == (int)TypeFormDoc.OfficeNote) doc = "Uzcard_2.xls";
                if (id_type == (int)TypeDoc.SendToFilial) doc = "Uzcard_3.xls";

                System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                ExcelAp ep = new ExcelAp();
                if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + doc))
                {
                    ep.SetWorkSheet(1);
                    ZapExcelUz(id_type, ep);

                    if (doc.Length > 0 && WebConfigurationManager.AppSettings["DocPath"] != null)
                    {
                        doc = String.Format("{0}Temp\\{1}", WebConfigurationManager.AppSettings["DocPath"], doc);
                        ep.SaveAsDoc(doc, false);
                        docTempName = doc;
                    }
                }
                ep.Close();
                GC.Collect();
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                if (doc.Length > 0)
                    ep.ReturnXls(Response, doc);
            }
            catch (Exception ex)
            {
                WebLog.LogClass.WriteToLog(ex.ToString());
            }
        }
        private void ZapExcel(int id_type, int spec_type, int id_deliv, ExcelAp ep)
        {
            string dt_doc = "";
            int all_cnt_n = 0; int all_cnt_p = 0; int all_cnt_b = 0;
            int cnt_prod = 0;
            int cnt = 0, cntb = 0;
            int type_p = 0;
            string prod_name = "";
            DataSet ds = new DataSet();
            #region Записка на выпуск карт - OfficeNote
            if (id_type == (int)TypeFormDoc.OfficeNote)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
                ep.SetText_Name("dt_doc", dt_doc);

                ds.Clear();
                res = ExecuteQuery(String.Format("select prefix_file, count(id) as cnt from V_Cards_StorageDocs where (id_stat = 1) and (id_doc={0}) group by prefix_file", id_doc.ToString()), ref ds, null);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    ep.SetText_Name(ds.Tables[0].Rows[i]["prefix_file"].ToString(), ds.Tables[0].Rows[i]["cnt"].ToString());
            }
            #endregion
            #region Записка на выпуск карт 1 - OfficeNote1
            if (id_type == (int)TypeFormDoc.OfficeNote1)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
                ep.SetText_Name("dt_doc", dt_doc);
                ds.Clear();
                res = ExecuteQuery(String.Format("select id_prb, id_prod, prod_name from V_ProductsBanks where parent is NULL order by id_sort"), ref ds, null);
                SqlCommand sel = new SqlCommand();
                sel.CommandText = "select count(id) from V_Cards_StorageDocs where (id_stat=1) and (id_doc=@doc) and (id_prb=@prb or id_prb in (select id_prb from V_ProductsBanks where parent=@prd))";
                sel.Parameters.Add("@doc", SqlDbType.Int).Value = id_doc;
                sel.Parameters.Add("@prb", SqlDbType.Int);
                sel.Parameters.Add("@prd", SqlDbType.Int);
                int col = 0;
                object obj = null;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    sel.Parameters["@prb"].Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                    sel.Parameters["@prd"].Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prod"]);
                    ExecuteScalar(sel, ref obj, null);
                    if (obj != null && (int)obj > 0)
                    {
                        ep.SetText(15, 1 + col, ds.Tables[0].Rows[i]["prod_name"].ToString());
                        ep.SetText(16, 1 + col, obj.ToString());
                        col++;
                    }
                }
                if (col > 0) ep.SetRangeBorders(15, 1, 16, col);
                sel.CommandText = "select count(id) from V_Cards_StorageDocs where (id_stat=1) and (id_doc=@doc)";
                ExecuteScalar(sel, ref obj, null);
                if (obj != null)
                    ep.SetText(18, 1, "Итого карт: " + obj.ToString());
                sel.CommandText = "select cnt_perso from V_Products_StorageDocs where (id_doc=@doc) and (id_type=@tp)";
                sel.Parameters.Add("@tp", SqlDbType.Int).Value = 2;
                ExecuteScalar(sel, ref obj, null);
                if (obj != null)
                    ep.SetText(19, 1, "Итого пин-конвертов: " + obj.ToString());

            }
            #endregion
            #region Сорт лист - SortList
            if (id_type == (int)TypeFormDoc.SortList)
            {

                for (int ii = 0; ii < 2; ii++)
                {
                    cnt_prod = 0;
                    string sql_dop = "";
                    if (ii == 0)
                    {
                        ep.SetWorkSheet(1);
                        ep.AddWorkSheet(1);
                        ep.SetWorkSheetName(1, "Карты");
                        ep.SetWorkSheet(1);
                    }
                    else
                    {
                        ep.SetWorkSheet(2);
                        ep.SetWorkSheetName(2, "Пин конверты");
                        sql_dop = " and (id in (select c.id from Cards c where id=c.id and c.isPin>0))"; // для подсчета пин-конвертов
                    }
                    dt_doc = String.Format("{0:dd.MM.yyyy HH:mm}", DateTime.Now);
                    ep.SetText(6, 5, dt_doc);

                    string dep_c = "";
                    string dep = "";
                    int cnt_all = 0; int cntf = 0; int cnt_allf = 0; int sum_all = 0;
                    int cur_row = 8;
                    int cnt_head = 0;
                    int cnt_dep = 0;

                    try
                    {
                        cnt_head = Convert.ToInt32(ConfigurationSettings.AppSettings["SortListH"]);
                    }
                    catch { }

                    int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);

                    //ds.Clear();
                    //res = ExecuteQuery("select distinct prod_name from V_SortList order by prod_name", ref ds, null);
                    //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    //{
                    //    ep.SetText(cur_row , 2+i , ds.Tables[0].Rows[i]["prod_name"].ToString());
                    //    cnt_prod++;
                    //}

                    ds.Clear();
                    res = ExecuteQuery(String.Format("select distinct prod_name,bin,prefix_ow,id_sort from V_Cards where (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) order by id_sort", id_doc.ToString()), ref ds, null);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        //ep.SetText(cur_row, 2 + i, ds.Tables[0].Rows[i]["prod_name"].ToString());
                        ep.SetText(cur_row, 2 + i, ds.Tables[0].Rows[i]["bin"].ToString() + "_" + ds.Tables[0].Rows[i]["prefix_ow"].ToString());
                        cnt_prod++;
                    }

                    ep.SetText(cur_row, 2 + cnt_prod, "Всего :");

                    ds.Clear();
                    //res = ExecuteQuery("select department, prod_name, cnt,cnt_all from V_SortList order by department, prod_name", ref ds, null);
                    //res = ExecuteQuery(String.Format("select DepBranchCard AS department,prod_name,id_prb,id_branchCard,bin,prefix_ow,(select Count(id_prb) as Expr1 from Cards where (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id_stat=1 or id_stat = 2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) as cnt,(select Count(id_prb) as Expr1 from V_Cards where (id_stat=1 or id_stat=2) and (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) or (id_stat=1 or id_stat=2) AND (id_prb=gg.id_prb) and (id_BranchCard_parent=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) as cnt_all from V_Cards gg where (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) group by DepBranchCard, prod_name, id_prb, id_branchCard,bin,prefix_ow order by department,prod_name",id_doc.ToString()), ref ds, null);
                    //res = ExecuteQuery(String.Format("select department,prod_name,bin,prefix_ow,(select count(id_prb) as cnt from Cards where (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) as cnt,(select count(id_prb) as cnt from V_Cards where (id_stat=1 or id_stat=2) and (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) or (id_stat=1 or id_stat=2) and (id_prb=gg.id_prb) and (id_BranchCard_parent=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) as cnt_all from (select DepBranchCard as department,prod_name,id_prb,id_branchCard,bin,prefix_ow from V_Cards gg where (id_stat=1 or id_stat=2) and (id in (select id_card from V_Cards_StorageDocs where id_doc={0})) union select DepBranchCard_parent as department,prod_name,id_prb,id_BranchCard_parent,bin,prefix_ow from V_Cards gg where (id_stat=1 or id_stat=2) and (id_BranchCard_parent in (select id_branchcard_parent from V_Cards_StorageDocs where id_doc={0} and id_branchcard_parent<>0)) group by DepBranchCard_parent,prod_name,id_prb,id_BranchCard_parent,bin,prefix_ow) gg order by department,prod_name", id_doc.ToString()), ref ds, null);
                    //res = ExecuteQuery(String.Format("select department,prod_name,bin,prefix_ow,(select count(id_prb) as cnt from Cards where (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) as cnt,(select count(id_prb) as cnt from Cards where (id_branchCard=gg.id_branchCard) and (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) AS cntf,(select count(id_prb) as cnt from V_Cards where (id_stat=1 or id_stat=2) and (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) or (id_stat=1 or id_stat=2) and (id_prb=gg.id_prb) and (id_BranchCard_parent=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) as cnt_all,(select count(id_prb) as cnt from V_Cards where (id_stat=1 or id_stat=2) and (id_branchCard=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) or (id_stat=1 or id_stat=2) and (id_BranchCard_parent=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) AS cnt_allf from (select DepBranchCard as department,prod_name,id_prb,id_branchCard,bin,prefix_ow,depbranchcardsort from V_Cards gg where (id_stat=1 or id_stat=2) and (id in (select id_card from V_Cards_StorageDocs where id_doc={0})) union select DepBranchCard_parent as department,prod_name,id_prb,id_BranchCard_parent,bin,prefix_ow,depbranchcardsort from V_Cards gg where (id_stat=1 or id_stat=2) and (id_BranchCard_parent in (select id_branchcard_parent from V_Cards_StorageDocs where id_doc={0} and id_branchcard_parent<>0)) group by DepBranchCard_parent,prod_name,id_prb,id_BranchCard_parent,bin,prefix_ow,depbranchcardsort) gg order by depbranchcardsort,department,prod_name", id_doc.ToString()), ref ds, null);
                    res = ExecuteQuery(String.Format("select department,prod_name,bin,prefix_ow,id_branchCard,id_prb," +
                        "(select id_parent from Branchs b where id_branchCard=b.id) as id_branchCard_parent," +
                        "(select count(b.id) from Branchs b where isolated!=0 and id_branchCard=b.id) as isolated," +
                        "(select count(id_prb) as cnt from Cards where (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0})){1}) as cnt," +
                        "(select count(id_prb) as cnt from Cards where (id_branchCard=gg.id_branchCard) and (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0})){1}) AS cntf," +
                        "(select count(id_prb) as cnt from V_Cards where (id_stat=1 or id_stat=2) and (id_prb=gg.id_prb) and (id_branchCard=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) or (id_stat=1 or id_stat=2) and (id_prb=gg.id_prb) and (id_BranchCard_parent=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})){1}) as cnt_all," +
                        "(select count(id_prb) as cnt from V_Cards where (id_stat=1 or id_stat=2) and (id_branchCard=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})) or (id_stat=1 or id_stat=2) and (id_BranchCard_parent=gg.id_branchCard) and (id in (select id_card from Cards_StorageDocs where id_doc={0})){1}) AS cnt_allf " +
                        " from (select DepBranchCard as department,prod_name,id_prb,id_branchCard,bin,prefix_ow,depbranchcardsort from V_Cards gg where (id_stat=1 or id_stat=2) and (id in (select id_card from V_Cards_StorageDocs where id_doc={0})) union select DepBranchCard_parent as department,prod_name,id_prb,id_BranchCard_parent,bin,prefix_ow,depbranchcardsort from V_Cards gg where (id_stat=1 or id_stat=2) and (id_BranchCard_parent in (select id_branchcard_parent from V_Cards_StorageDocs2 where id_doc={0} and id_branchcard_parent<>0)) group by DepBranchCard_parent,prod_name,id_prb,id_BranchCard_parent,bin,prefix_ow,depbranchcardsort) gg order by depbranchcardsort,department,prod_name", id_doc.ToString(), sql_dop), ref ds, null);

                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        //----------------------------------
                        // Корректировка цифр для филиалов с isolated==true
                        //----------------------------------
                        int id_branchCard = Convert.ToInt32(ds.Tables[0].Rows[i]["id_branchCard"]);
                        int id_branchCard_parent = Convert.ToInt32(ds.Tables[0].Rows[i]["id_branchCard_parent"]);
                        if (id_branchCard_parent == 0 && id_branchCard > 0)
                        {
                            int sum = 0;
                            for (int j = i; j < ds.Tables[0].Rows.Count; j++)
                            {
                                if (Convert.ToInt32(ds.Tables[0].Rows[j]["id_branchCard_parent"]) == id_branchCard_parent && Convert.ToInt32(ds.Tables[0].Rows[j]["id_branchCard"]) == id_branchCard)
                                    continue;
                                if (Convert.ToInt32(ds.Tables[0].Rows[j]["id_branchCard_parent"]) != id_branchCard)
                                {
                                    break;
                                }
                                if (Convert.ToInt32(ds.Tables[0].Rows[j]["isolated"]) != 0)
                                {
                                    int cnt_al = Convert.ToInt32(ds.Tables[0].Rows[j]["cnt"]);
                                    int id_prb = Convert.ToInt32(ds.Tables[0].Rows[j]["id_prb"]);
                                    sum += cnt_al;
                                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_branchCard_parent"]) == 0 &&
                                            Convert.ToInt32(ds.Tables[0].Rows[j]["id_branchCard_parent"]) == Convert.ToInt32(ds.Tables[0].Rows[i]["id_branchCard"]) &&
                                                id_prb == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                                    {
                                        int c_al = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_all"]);
                                        c_al -= cnt_al;
                                        ds.Tables[0].Rows[i]["cnt_all"] = c_al;
                                    }

                                }
                            }
                            if (sum > 0)
                            {
                                int c_allf = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_allf"]);
                                c_allf -= sum;
                                ds.Tables[0].Rows[i]["cnt_allf"] = c_allf;
                            }
                        }
                    }

                    // это бред, зачем цифры корректировал перед этим??????
                    /*
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        if (Convert.ToInt32(ds.Tables[0].Rows[i]["isolated"]) != 0)
                        {
                            ds.Tables[0].Rows.RemoveAt(i);
                            i--;
                        }
                    }
                    */

                    //----------------------------------
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        //dep = ds.Tables[0].Rows[i]["department"].ToString();
                        dep = ds.Tables[0].Rows[i]["department"].ToString();
                        int tt = 0;
                        for (tt = 0; tt < dep.Length; tt++)
                        {
                            if (Char.IsLetter(dep[tt]))
                                break;
                        }
                        if (tt > 0)
                            dep = dep.Substring(0, tt);

                        //prod_name = ds.Tables[0].Rows[i]["prod_name"].ToString();
                        prod_name = ds.Tables[0].Rows[i]["bin"].ToString() + "_" + ds.Tables[0].Rows[i]["prefix_ow"].ToString();
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                        cnt_all = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_all"]);
                        cntf = Convert.ToInt32(ds.Tables[0].Rows[i]["cntf"]);
                        cnt_allf = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_allf"]);

                        if (dep_c != dep)
                        {

                            cnt_dep++;
                            //повторение шапки
                            if (cnt_head != 0 && cnt_dep % cnt_head == 0)
                            {
                                cur_row++;
                                for (int k = 0; k < cnt_prod; k++)
                                    ep.SetText(cur_row, 2 + k, ep.GetCell(8, 2 + k));
                                ep.SetText(cur_row, 2 + cnt_prod, "Всего :");
                            }
                            cur_row++;
                            ep.SetText(cur_row, 1, dep);
                            if (cntf == cnt_allf)
                                ep.SetText(cur_row, 2 + cnt_prod, cntf.ToString());
                            else
                                ep.SetText(cur_row, 2 + cnt_prod, cntf.ToString() + " (" + cnt_allf.ToString() + ")");
                            sum_all = sum_all + cntf;
                        }
                        for (int k = 0; k < cnt_prod; k++)
                        {
                            if (ep.GetCell(8, 2 + k) == prod_name)
                            {
                                if (cnt == cnt_all)
                                    ep.SetText(cur_row, 2 + k, cnt.ToString());
                                else
                                    ep.SetText(cur_row, 2 + k, cnt.ToString() + " (" + cnt_all.ToString() + ")");
                                break;
                            }
                        }
                        //                    dep_c = ds.Tables[0].Rows[i]["department"].ToString();
                        dep_c = dep;
                    }
                    cur_row++;
                    ep.SetText(cur_row, 1, "Всего :");

                    ds.Clear();
                    //res = ExecuteQuery("select prod_name, Sum(cnt) as cnt from V_SortList group by prod_name", ref ds, null);

                    /*//изменения сделанные в Казани для ускорения
                    res = ExecuteQuery(String.Format("select prod_name,bin,prefix_ow,sum(cnt) AS cnt from (select distinct prod_name,bin,prefix_ow,(select Count(id_prb) as Expr1 from Cards where (id_prb=gg.id_prb) and (id_stat= 1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) as cnt from V_Cards gg where (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))) DERIVEDTBL group by prod_name,bin,prefix_ow",id_doc.ToString()), ref ds, null);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        //prod_name = ds.Tables[0].Rows[i]["prod_name"].ToString();
                        prod_name = ds.Tables[0].Rows[i]["bin"].ToString() + "_" + ds.Tables[0].Rows[i]["prefix_ow"].ToString();
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);

                        for (int k = 0; k < cnt_prod; k++)
                        {
                            if (ep.GetCell(8, 2 + k) == prod_name)
                            {
                                ep.SetText(cur_row, 2 + k, cnt.ToString());
                                break;
                            }
                        }
                    }
                    */
                    res = ExecuteQuery(String.Format("select distinct id_prb,prod_name,bin,prefix_ow from V_Cards gg where (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))", id_doc), ref ds, null);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        prod_name = ds.Tables[0].Rows[i]["bin"].ToString() + "_" + ds.Tables[0].Rows[i]["prefix_ow"].ToString();
                        int id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                        object obj = null;
                        ExecuteScalar(String.Format("select Count(id_prb) from Cards where (id_prb={0}) and (id_stat= 1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={1}){2})", id_prb, id_doc, sql_dop), ref obj, null);
                        for (int k = 0; k < cnt_prod; k++)
                        {
                            if (ep.GetCell(8, 2 + k) == prod_name)
                            {
                                ep.SetText(cur_row, 2 + k, obj.ToString());
                                break;
                            }
                        }
                    }
                    //   ds.Clear();
                    //   res = ExecuteQuery(String.Format("select count(id_prb) as cnt from Cards where (id_stat=1 or id_stat=2) and (id in (select id_card from Cards_StorageDocs where id_doc={0}))", id_doc.ToString()), ref ds, null);
                    //   cnt = Convert.ToInt32(ds.Tables[0].Rows[0]["cnt"]);

                    ep.SetText(cur_row, 2 + cnt_prod, sum_all.ToString());

                    ep.SetRangeBorders(8, 1, cur_row, 2 + cnt_prod);
                    ep.SetRangeAutoFit(8, 1, cur_row, 2 + cnt_prod);
                    // ep.SetRangeBold(8, 1, 8, 1 + cnt_prod);
                    // ep.SetRangeBold(8, 1, 8+cur_row, 1);
                    ep.SetRangeAlignment(8, 2, cur_row, 2 + cnt_prod, Excel.Constants.xlCenter);
                }

            }
            #endregion
            #region Приложение2 - CardFromStorage
            if (id_type == (int)TypeDoc.CardFromStorage)
            {
                dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
                ep.SetText(10, 1, "Дата: " + dt_doc);
                ep.SetText(49, 1, dt_doc);
                ep.SetText(49, 6, dt_doc);

                for (int j = 0; j < gvProducts.Rows.Count; j++)
                {
                    prod_name = gvProducts.DataKeys[j].Values["prod_name"].ToString();
                    type_p = Convert.ToInt32(gvProducts.DataKeys[j].Values["id_type"]);

                    if (type_p == 1)
                    {
                        if (spec_type == 8)
                            cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_perso"]) + Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_brak"]);
                        else
                            cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_new"]);

                        ep.SetText(cnt_prod + 13, 1, "  " + (j + 1).ToString() + ".   " + prod_name);
                        ep.SetText(cnt_prod + 13, 6, cnt.ToString());
                        all_cnt_n = all_cnt_n + cnt;
                        cnt_prod = cnt_prod + 1;
                    }
                }
                ep.SetText(cnt_prod + 13, 1, "   Всего");
                ep.SetText(cnt_prod + 13, 6, all_cnt_n.ToString());
            }
            #endregion
            #region Приложение3 - PinFromStorage
            if (id_type == (int)TypeDoc.PinFromStorage)
            {
                dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
                ep.SetText(12, 1, "Дата: " + dt_doc);
                ep.SetText(47, 1, dt_doc);
                ep.SetText(47, 6, dt_doc);

                for (int j = 0; j < gvProducts.Rows.Count; j++)
                {
                    type_p = Convert.ToInt32(gvProducts.DataKeys[j].Values["id_type"]);
                    cnt = 0; cntb = 0;
                    if (type_p == 2)
                    {
                        if (spec_type == 8)
                        {
                            cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_perso"]);
                            cntb = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_brak"]);
                        }
                        else
                            cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_new"]);
                    }

                    all_cnt_n = all_cnt_n + cnt + cntb;
                }
                ep.SetText(22, 8, all_cnt_n.ToString());
            }
            #endregion
            #region Приложение4 - CardToStorage
            if (id_type == (int)TypeDoc.CardToStorage)
            {
                dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
                ep.SetText(10, 1, "Дата: " + dt_doc);
                ep.SetText(49, 1, dt_doc);
                ep.SetText(49, 6, dt_doc);

                for (int j = 0; j < gvProducts.Rows.Count; j++)
                {
                    prod_name = gvProducts.DataKeys[j].Values["prod_name"].ToString();
                    type_p = Convert.ToInt32(gvProducts.DataKeys[j].Values["id_type"]);

                    if (type_p == 1)
                    {
                        ep.SetText(cnt_prod + 13, 1, "  " + (cnt_prod + 1).ToString() + ".   " + prod_name);

                        if (spec_type != 8)
                        {
                            cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_new"]);
                            ep.SetText(cnt_prod + 13, 8, cnt.ToString());
                            all_cnt_n = all_cnt_n + cnt;
                        }

                        cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_perso"]);
                        ep.SetText(cnt_prod + 13, 9, cnt.ToString());
                        all_cnt_p = all_cnt_p + cnt;

                        cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_brak"]);
                        ep.SetText(cnt_prod + 13, 10, cnt.ToString());
                        all_cnt_b = all_cnt_b + cnt;

                        cnt_prod = cnt_prod + 1;
                    }
                }
                ep.SetText(cnt_prod + 13, 1, "   Всего");
                if (spec_type != 8)
                    ep.SetText(cnt_prod + 13, 8, all_cnt_n.ToString());
                ep.SetText(cnt_prod + 13, 9, all_cnt_p.ToString());
                ep.SetText(cnt_prod + 13, 10, all_cnt_b.ToString());
            }
            #endregion
            #region Приложение5 - PinToStorage
            if (id_type == (int)TypeDoc.PinToStorage)
            {
                dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
                ep.SetText(12, 1, "Дата: " + dt_doc);
                ep.SetText(47, 1, dt_doc);
                ep.SetText(47, 6, dt_doc);

                for (int j = 0; j < gvProducts.Rows.Count; j++)
                {
                    type_p = Convert.ToInt32(gvProducts.DataKeys[j].Values["id_type"]);

                    if (type_p == 2)
                    {
                        if (spec_type != 8)
                        {
                            cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_new"]);
                            all_cnt_n = all_cnt_n + cnt;
                        }

                        cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_perso"]);
                        all_cnt_p = all_cnt_p + cnt;

                        cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_brak"]);
                        all_cnt_b = all_cnt_b + cnt;
                    }
                }
                if (spec_type != 8)
                    ep.SetText(22, 7, all_cnt_n.ToString());
                ep.SetText(22, 8, all_cnt_p.ToString());
                ep.SetText(22, 9, all_cnt_b.ToString());

            }
            #endregion
            #region Приложение6 - SendToFilial (id_deliv=0)
            if ((id_type == (int)TypeDoc.SendToFilial) && (id_deliv == 0))
            {
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region Приложение6 прием в филиале - ReceiveToFilial
            if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToFilialPacket)
            {
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region Приложение6 отказ в приеме в филиале - DontReceiveToFilial
            if (id_type == (int)TypeDoc.DontReceiveToFilial)
            {
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region Приложение6 - 2 акта + реестр - Akt2Reestr (id_deliv=0)
            if ((id_type == (int)TypeFormDoc.Akt2Reestr) && (id_deliv == 0))
            {
                ep.SetWorkSheet(1);
                ep.AddWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString() + "_1");
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                ep.SetWorkSheet(3);
                ep.SetWorkSheetName(3, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region Приложение6_all - SendToFilial (id_deliv > 0)
            if ((id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeFormDoc.Akt2Reestr) && (id_deliv > 0))
            {
                //вывод многими листами
                /*
                ds.Clear();
                
                res = ExecuteQuery(String.Format("select id,id_branch,branch from V_StorageDocs where id_deliv={0} order by branch", id_deliv.ToString()), ref ds, null);
                
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                        ep.SetWorkSheet(1);
                        if (i != ds.Tables[0].Rows.Count - 1) ep.AddWorkSheet();
                        ep.SetWorkSheet(2);
                        if (i != ds.Tables[0].Rows.Count - 1) ep.AddWorkSheet();
                }
                
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetWorkSheet(2 * i + 1);
                    ep.SetWorkSheetName(2 * i + 1, "a_" + ds.Tables[0].Rows[i]["branch"].ToString());
                    ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]), ds.Tables[0].Rows[i]["branch"].ToString(), ep);
                    ep.SetWorkSheet(2 * i + 2);
                    ep.SetWorkSheetName(2 * i + 2, "r_" + ds.Tables[0].Rows[i]["branch"].ToString());
                    ExcelReestrFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), ep);
                }
                */
                ds.Clear();
                res = ExecuteQuery(String.Format("select id,id_branch,branch from V_StorageDocs where id_deliv={0} order by branch", id_deliv.ToString()), ref ds, null);
                int countBranchs = 0;


                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    int[] branchs = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), false);
                    int[] branchsIsolated = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), true);
                    if (branchs.Length > 0) countBranchs++;
                    if (branchsIsolated.Length > 0) countBranchs += branchsIsolated.Length;
                }

                for (int i = 0; i < countBranchs; i++)
                {
                    ep.SetWorkSheet(1);
                    if (i != countBranchs - 1) ep.AddWorkSheet();
                    ep.SetWorkSheet(2);
                    if (i != countBranchs - 1) ep.AddWorkSheet();
                }

                for (int i = 0, n = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    int[] branchs = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), false);
                    BranchStore[] branchsIsolated = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]));
                    if (branchs.Length > 0)
                    {
                        ep.SetWorkSheet(2 * n + 1);
                        ep.SetWorkSheetName(2 * n + 1, "a_" + ds.Tables[0].Rows[i]["branch"].ToString());
                        ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]), ds.Tables[0].Rows[i]["branch"].ToString(), ep, branchsIsolated, null);
                        ep.SetWorkSheet(2 * n + 2);
                        ep.SetWorkSheetName(2 * n + 2, "r_" + ds.Tables[0].Rows[i]["branch"].ToString());
                        ExcelReestrFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), ep, branchsIsolated, null);
                        n++;
                    }
                    if (branchsIsolated.Length > 0)
                    {
                        for (int j = 0; j < branchsIsolated.Length; j++)
                        {
                            ep.SetWorkSheet(2 * n + 1);
                            ep.SetWorkSheetName(2 * n + 1, "a_" + branchsIsolated[j].department);
                            ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), branchsIsolated[j].id, branchsIsolated[j].department, ep, null, branchsIsolated[j]);
                            ep.SetWorkSheet(2 * n + 2);
                            ep.SetWorkSheetName(2 * n + 2, "r_" + branchsIsolated[j].department);
                            ExcelReestrFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), ep, null, branchsIsolated[j]);
                            n++;
                        }
                    }
                }
                if (id_type == (int)TypeFormDoc.Akt2Reestr)
                {
                    for (int i = 0; i < countBranchs; i++)
                    {
                        ep.SetWorkSheet(2 * i + 1 + i);
                        ep.AddWorkSheet(2 * i + 1 + i);
                    }
                }
            }
            #endregion
            #region Приложение6_all c двумя актами и реестром - Akt2Reestr (id_deliv > 0)
            /* Перенес в Приложение6_all - SendToFilial (id_deliv > 0) */
            /*if (id_type == (int)TypeFormDoc.Akt2Reestr && (id_deliv > 0))
            {
                //вывод многими листами
                ds.Clear();
                res = ExecuteQuery(String.Format("select id,id_branch,branch from V_StorageDocs where id_deliv={0} order by branch", id_deliv.ToString()), ref ds, null);
                ep.SetWorkSheet(1);
                ep.AddWorkSheet(1);
                for (int i = 0; i < ds.Tables[0].Rows.Count - 1; i++)
                {
                    ep.SetWorkSheet(1);
                    ep.AddWorkSheet();
                    ep.AddWorkSheet();
                    ep.SetWorkSheet(3);
                    ep.AddWorkSheet();
                }

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetWorkSheet(3 * i + 1);
                    ep.SetWorkSheetName(3 * i + 1, "a_" + ds.Tables[0].Rows[i]["branch"].ToString());
                    ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]), ds.Tables[0].Rows[i]["branch"].ToString(), ep);
                    ep.SetWorkSheet(3 * i + 2);
                    ep.SetWorkSheetName(3 * i + 2, "a_" + ds.Tables[0].Rows[i]["branch"].ToString() + "_1");                  
                    ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]), ds.Tables[0].Rows[i]["branch"].ToString(), ep);
                    ep.SetWorkSheet(3 * i + 3);
                    ep.SetWorkSheetName(3 * i + 3, "r_" + ds.Tables[0].Rows[i]["branch"].ToString());
                    ExcelReestrFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), ep);
                }
            }
            */
            #endregion
            #region Приложение6_all с двумя актами и реестром для персонализации - Akt2ReestrPerso
            if (id_type == (int)TypeFormDoc.Akt2ReestrPerso)
            {
                // получаем список продуктов

                ds.Clear();
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                res = (string)ExecuteCommand(comm, ref ds, null);
                ArrayList al = new ArrayList();
                foreach (DataRow dr in ds.Tables[0].Rows)
                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                ds.Clear();
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                res = ExecuteQuery(String.Format("select distinct id_branchcard, BranchCard from V_cards_StorageDocs where id_doc={0} order by branchcard", id_doc), ref ds, null);
                ep.SetWorkSheet(1);
                ep.AddWorkSheet(1);
                for (int i = 0; i < ds.Tables[0].Rows.Count - 1; i++)
                {
                    ep.SetWorkSheet(1);
                    ep.AddWorkSheet();
                    ep.AddWorkSheet();
                    ep.SetWorkSheet(3);
                    ep.AddWorkSheet();
                }
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    foreach (MyProd mp in al)
                        for (int t = 0; t < mp.cnts.Length; t++)
                            mp.cnts[t] = 0;
                    ep.SetWorkSheetName(3 * i + 1, "a_" + ds.Tables[0].Rows[i]["branchcard"].ToString());
                    ep.SetWorkSheetName(3 * i + 2, "a_" + ds.Tables[0].Rows[i]["branchcard"].ToString() + "_1");
                    ep.SetWorkSheetName(3 * i + 3, "r_" + ds.Tables[0].Rows[i]["branchcard"].ToString());
                    Excel2AktReestrPerso(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_branchcard"]), Convert.ToString(ds.Tables[0].Rows[i]["branchcard"]), 3 * i, al, ep);
                }
            }
            #endregion
            #region отправка филиал-филиал без рассылки - SendToFilialFilial (id_deliv = 0)
            if (id_type == (int)TypeDoc.SendToFilialFilial && id_deliv == 0)
            {
                string brname = gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString();
                if (brname.IndexOf("-->") >= 0)
                    brname = brname.Split(new string[] { "-->" }, StringSplitOptions.None)[1];
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + brname);
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), id_type, ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + brname);
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region отправка филиал-филиал с рассылки - SendToFilialFilial (id_deliv > 0)
            if (id_type == (int)TypeDoc.SendToFilialFilial && id_deliv > 0)
            {
                ds.Clear();
                res = ExecuteQuery(String.Format("select id,id_branch,branch from V_StorageDocs where id_deliv={0} order by branch", id_deliv.ToString()), ref ds, null);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetWorkSheet(1);
                    if (i != ds.Tables[0].Rows.Count - 1) ep.AddWorkSheet();
                    ep.SetWorkSheet(2);
                    if (i != ds.Tables[0].Rows.Count - 1) ep.AddWorkSheet();
                }
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetWorkSheet(2 * i + 1);
                    ep.SetWorkSheetName(2 * i + 1, "a_" + ds.Tables[0].Rows[i]["branch"].ToString());
                    ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]), ds.Tables[0].Rows[i]["branch"].ToString(), id_type, ep);
                    ep.SetWorkSheet(2 * i + 2);
                    ep.SetWorkSheetName(2 * i + 2, "r_" + ds.Tables[0].Rows[i]["branch"].ToString());
                    ExcelReestrFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), ep);
                }
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetWorkSheet(2 * i + 1 + i);
                    ep.AddWorkSheet(2 * i + 1 + i);
                    //ep.SetWorkSheetName(2 * i + i, "a_" + ds.Tables[0].Rows[i]["branch"].ToString() + "_1");

                }
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetWorkSheet(3 * i + 2 + i);
                    ep.AddWorkSheet(3 * i + 2 + i);

                }
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetWorkSheet(4 * i + 3 + i);
                    ep.AddWorkSheet(4 * i + 3 + i);

                }


            }
            #endregion
            #region акт приема-передачи для транспортной компании - TransportAct
            if (id_type == (int)TypeFormDoc.TransportAct)
            {
                ds.Clear();
                res = ExecuteQuery(String.Format("select id,id_branch,branch from V_StorageDocs where id_deliv={0} order by branch", id_deliv.ToString()), ref ds, null);
                ep.SetWorkSheet(1);
                cnt = 0;
                int rw = 5;
                ep.SetText_Name("DateForm", String.Format("{0:dd.MM.yyyy}", Convert.ToDateTime(gvDocs.DataKeys[gvCards.SelectedIndex].Values["date_doc"])));
                ep.SetText_Name("DateForm1", String.Format("{0:dd MMMM yyyy}", DateTime.Now));
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    object obj = null, obj1 = null;
                    string[] strs = ConfigurationSettings.AppSettings["Transport"].Split(',');
                    for (int t = 0; t < strs.Length; t++)
                        strs[t] = String.Format("prefix_ow like '{0}'", strs[t]);
                    res = ExecuteScalar(String.Format("select count(*) from V_Cards_StorageDocs where id_doc={0} and ({1})", ds.Tables[0].Rows[i]["id"], String.Join(" or ", strs)), ref obj, null);
                    if ((int)obj > 0)
                    {
                        ep.SetText(rw, 1, ds.Tables[0].Rows[i]["branch"].ToString());
                        res = ExecuteScalar(String.Format("select count(*) from V_Products_StorageDocs where id_doc={0} and id_type=2", ds.Tables[0].Rows[i]["id"]), ref obj1, null);
                        ep.SetText(rw, 2, String.Format("{0}", obj));
                        cnt += (int)obj;
                        rw++;
                    }
                }
                ep.ShowRows(5, rw);
                ep.SetText_Name("Itogo", cnt.ToString());
            }
            #endregion
            #region акт приема-передачи для почты россии - PochtaAct
            if (id_type == (int)TypeFormDoc.PochtaAct)
            {
                ds.Clear();
                res = ExecuteQuery(String.Format("select branch from V_StorageDocs where id_deliv={0} order by branch", id_deliv.ToString()), ref ds, null);
                ep.SetWorkSheet(1);
                ep.SetText_Name("DateF", String.Format("{0:dd.MM.yyyy}", Convert.ToDateTime(gvDocs.DataKeys[gvCards.SelectedIndex].Values["date_doc"])));
                int rw = 8;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetText(rw, 1, (i + 1).ToString());
                    ep.SetText(rw, 2, ds.Tables[0].Rows[i]["branch"].ToString());
                    ep.SetText(rw, 7, "1");
                    rw++;
                }
                ep.SetText_Name("Itogo", ds.Tables[0].Rows.Count.ToString());
                ep.ShowRows(8, rw);
            }
            #endregion
            #region акт на уничтожение карт - DeleteBrak
            if (id_type == (int)TypeDoc.DeleteBrak)
            {
                ds.Clear();
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                res = (string)ExecuteCommand(comm, ref ds, null);
                ArrayList al = new ArrayList();
                foreach (DataRow dr in ds.Tables[0].Rows)
                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                ds.Clear();
                comm.CommandText = "select id_prb, cnt_brak from V_Rep_Moving where type=11 and (id_type=1 or id_type=2) and doc_id=" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"].ToString();
                ExecuteCommand(comm, ref ds, null);
                if (ds != null && ds.Tables.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        foreach (MyProd mp in al)
                        {
                            if (mp.Type != 1 && mp.Type != 2)
                                continue;
                            if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                                mp.cnts[0] += Convert.ToInt32(dr["cnt_brak"]);
                        }
                    }
                }
                // уничтожено карт не востребованных и с истекшим сроком годности
                object obj = 0;
                ds.Clear();
                comm.CommandText = "select id_prb, id_prop, isPin from V_Cards_StorageDocs where id_doc=" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"].ToString();
                ExecuteCommand(comm, ref ds, null);
                if (ds != null && ds.Tables.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        bool isexpired = false;
                        bool isnotget = false;
                        foreach (MyProd mp in al)
                        {
                            if (mp.Type != 1)
                                continue;
                            if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                            {
                                // истекший срок годности
                                if (Convert.ToInt32(dr["id_prop"]) == 8)
                                {
                                    mp.cnts[1]++; mp.cnts[0]--;
                                    isexpired = true;
                                }
                                // не востребовано
                                if (Convert.ToInt32(dr["id_prop"]) == 7)
                                {
                                    mp.cnts[2]++; mp.cnts[0]--;
                                    isnotget = true;
                                }
                            }
                        }

                        if (Convert.ToInt32(dr["isPin"]) != 1)
                        {
                            isexpired = false;
                            isnotget = false;
                        }
                        //перекидываем пины в нужную колонку
                        foreach (MyProd mp in al)
                        {
                            if (mp.Type != 2)
                                continue;
                            if (isexpired)
                            {
                                mp.cnts[1]++;
                                mp.cnts[0]--;
                            }
                            if (isnotget)
                            {
                                mp.cnts[2]++;
                                mp.cnts[0]--;
                            }

                        }
                    }
                }
                int t = 0, cnt1 = 0, cnt2 = 0, cnt3 = 0;
                DateTime ddt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);
                ep.SetText(6, 2, String.Format("составили настоящий Акт в том, что {0:dd.MM.yyyy} года в нашем присутствии в хранилище бюро персонализации путем измельчения уничтожены следующие материальные ценности:", ddt));
                foreach (MyProd mp in al)
                {
                    if (mp.Type == 1)
                    {
                        if ((mp.cnts[0] + mp.cnts[1] + mp.cnts[2]) > 0)
                        {
                            ep.SetText(9 + t, 2, String.Format("{0} ({1})", mp.Name, mp.Bank));
                            if (mp.cnts[0] > 0)
                            {
                                ep.SetText(9 + t, 3, String.Format("{0}шт", mp.cnts[0]));
                                cnt1 += mp.cnts[0];
                            }
                            if (mp.cnts[1] > 0)
                            {
                                ep.SetText(9 + t, 4, String.Format("{0}шт", mp.cnts[1]));
                                cnt2 += mp.cnts[1];
                            }
                            if (mp.cnts[2] > 0)
                            {
                                ep.SetText(9 + t, 5, String.Format("{0}шт", mp.cnts[2]));
                                cnt3 += mp.cnts[2];
                            }
                            t++;
                        }
                    }
                }
                ep.SetText(73, 3, String.Format("{0}шт", cnt1));
                ep.SetText(73, 4, String.Format("{0}шт", cnt2));
                ep.SetText(73, 5, String.Format("{0}шт", cnt3));
                ep.ShowRows(9, 9 + t);
                // пин-конверты
                foreach (MyProd mp in al)
                {
                    if (mp.Type == 2)
                    {
                        if ((mp.cnts[0] + mp.cnts[1] + mp.cnts[2]) > 0)
                        {
                            if (mp.cnts[0] > 0)
                                ep.SetText(75, 3, String.Format("{0}шт", mp.cnts[0]));
                            if (mp.cnts[1] > 0)
                                ep.SetText(75, 4, String.Format("{0}шт", mp.cnts[1]));
                            if (mp.cnts[2] > 0)
                                ep.SetText(75, 5, String.Format("{0}шт", mp.cnts[2]));

                            ep.ShowRows(74, 75);
                        }
                    }
                }
            }
            #endregion
            #region акт выдачи клиенту/организации - SendToClient
            if (id_type == (int)TypeDoc.SendToClient || id_type == (int)TypeDoc.SendToClientFromPodotchet)                
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                DateTime dt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);

                ds.Clear();
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                res = (string)ExecuteCommand(comm, ref ds, null);
                ArrayList al = new ArrayList();
                foreach (DataRow dr in ds.Tables[0].Rows)
                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                ds.Clear();
                comm.CommandText = "select id_prb, company from V_Cards_StorageDocs1 where id=@id";
                comm.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                ExecuteCommand(comm, ref ds, null);
                if (ds == null || ds.Tables.Count == 0)
                    return;
                int pin = 0;
                string company = "";
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if (dr["id_prb"] == DBNull.Value)
                        continue;
                    foreach (MyProd mp in al)
                    {
                        if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                        {
                            mp.cnts[0]++; pin++;
                            break;
                        }
                    }
                    if (dr["company"].ToString().Trim().Length > 1) // чтобы не брать физиков, будем считать что названий компаний из одной буквы нет
                        company = dr["company"].ToString().Trim();
                }                
                int i = 0;
                foreach (MyProd mp in al)
                {
                    if (mp.cnts[0] != 0)
                    {
                        ep.SetText(10 + i, 1, mp.Name);
                        ep.SetText(10 + i, 6, mp.cnts[0].ToString());
                        i++;
                    }
                }
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ep.SetText(82, 6, pin.ToString());
                ep.ShowRows(10,10+i);
                ep.SetText_Name("date_string1", String.Format("{0:dd.MM.yyyy}", dt));
                ep.SetText_Name("date_string2", String.Format("{0:dd.MM.yyyy}", dt));
                ep.SetText_Name("Organization", "");
                ep.SetText_Name("UPosition", sc.UserPosition(User.Identity.Name));
                ep.SetText_Name("UFio", String.Format("____________/ {0} /", sc.UserFIO(User.Identity.Name)));
                ep.SetText_Name("Warrent", "");
                ep.SetText_Name("WFio", String.Format("____________/________________"));

                object obj = null;
                ExecuteScalar("select department from Branchs where id=" + sc.BranchId(User.Identity.Name), ref obj, null);
                if (obj != null) ep.AddText(90, 1, " (" + obj.ToString() + ")");
                
                if (company.Length > 0)
                {
                    ds.Clear();
                    comm.Parameters.Clear();
                    if (Request.Form["resde"].ToString() != "")
                        comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where idP=" + Request.Form["resde"].ToString();
                    else
                    {
                        //comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where embosstitle='" + company + "'";
                        if (branch_main_filial > 0)
                        {
                            comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where embosstitle='" + company + "' and BranchMainFilialId=" + branch_main_filial.ToString();
                        }
                        else
                        {
                            comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where embosstitle='" + company + "' and BranchMainFilialId is null";
                        }
                    }
                    ExecuteCommand(comm, ref ds, null);
                    if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                    {
                    }
                    else
                    {
                        ep.SetText_Name("Organization", ds.Tables[0].Rows[0]["title"].ToString().Trim());
                        ep.SetText_Name("Warrent", String.Format("{0} {5} {1} паспорт {2} от {3:dd.MM.yyyy} выдан {4}", ds.Tables[0].Rows[0]["position"].ToString().Trim(), ds.Tables[0].Rows[0]["person"].ToString().Trim(), ds.Tables[0].Rows[0]["passport"].ToString().Trim(), Convert.ToDateTime(ds.Tables[0].Rows[0]["pdate"]), ds.Tables[0].Rows[0]["pdivision"].ToString().Trim(), ds.Tables[0].Rows[0]["title"].ToString().Trim()));
                        ep.SetText_Name("WFio", String.Format("____________/{0}", ds.Tables[0].Rows[0]["person"].ToString().Trim()));
                    }
                }
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region Филиал-Филиал (приложение 18) - FilialFilial
            if (id_type == (int)TypeDoc.FilialFilial)
            {
                SqlCommand comm = conn.CreateCommand();
                comm.CommandText = "select department from branchs where id="+gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_courier"].ToString();
                //DataSet ds = new DataSet();
                ds.Clear();
                object temp = ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ep.SetWorkSheet(1);
                    ep.SetWorkSheetName(1, "a_" + ds.Tables[0].Rows[0][0].ToString());
                    //ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_courier"]), ds.Tables[0].Rows[0][0].ToString(), ep);
                    ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_courier"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                    ep.SetWorkSheet(2);
                    ep.SetWorkSheetName(2, "r_" + ds.Tables[0].Rows[0][0].ToString());
                    ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
                }
            }
            #endregion
            #region Возврат брака в центр (приложение 17) - SendToBank
            if (id_type == (int)TypeDoc.SendToBank)
            {
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReturnFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region Выдача расходных материалов (приложение 19) - Expendables
            if (id_type == (int)TypeDoc.Expendables)
            {
                SqlCommand comm = conn.CreateCommand();
                comm.CommandText = String.Format("select * from V_Products_StorageDocs where id_doc={0} order by id_sort", Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]));
                ds.Clear();
                object temp = ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    int t=13;
                    ep.SetWorkSheet(1);
                    foreach(DataRow dr in ds.Tables[0].Rows)
                    {
                        ep.SetText(t, 1, dr["prod_name"].ToString());
                        ep.SetText(t, 6, dr["cnt_new"].ToString());
                        t++;
                    }
                    ep.ShowRows(13, t);
                    DateTime ddt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);
                    ep.SetText_Name("date1", ddt.ToShortDateString());
                    ep.SetText_Name("date2", ddt.ToShortDateString());
                }
            }
            #endregion
            #region Выдача карт и т.д на рекламные цели - Reklama
            if (id_type == (int)TypeDoc.Reklama)
            {
                SqlCommand comm = conn.CreateCommand();
                comm.CommandText = String.Format("select * from V_Products_StorageDocs where id_doc={0} order by id_sort", Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]));
                ds.Clear();
                object temp = ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    cnt = 0;
                    int t = 13;
                    ep.SetWorkSheet(1);
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        ep.SetText(t, 1, dr["prod_name"].ToString());
                        ep.SetText(t, 6, dr["cnt_new"].ToString());
                        if (Convert.ToInt32(dr["id_type"]) == 1)
                            cnt += Convert.ToInt32(dr["cnt_new"]);
                        t++;
                    }
                    ep.SetText(46, 6, cnt.ToString());
                    ep.ShowRows(13, t);
                    DateTime ddt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);
                    ep.SetText_Name("date1", ddt.ToShortDateString());
                    ep.SetText_Name("date2", ddt.ToShortDateString());
                }
            }
            #endregion
            #region отправка карт на упаковку - ToWrapping
            if (id_type == (int)TypeDoc.ToWrapping)
            {
                ep.SetWorkSheet(1);
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), -1, "", ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "reestr");
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region прием карт из упаковки - FromWrapping
            if (id_type == (int)TypeDoc.FromWrapping)
            {
                ep.SetWorkSheet(1);
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), -1, "", ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "reestr");
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region Сопроводительный документ
            if (id_type == (int)TypeFormDoc.AccompDoc && id_deliv > 0)
            {
                ep.SetWorkSheet(1);
                ds.Clear();
                //object obj = null;
                //ExecuteScalar("select date_doc from StorageDocs where id_deliv=" + id_deliv.ToString(), ref obj, null);
                //try
                //{
                //    ep.SetText(3, 1, "Дата: " + String.Format("{0:dd.MM.yyyy}", Convert.ToDateTime(obj)));
                //}
                //catch { }
                //ExecuteQuery(String.Format(@"select id, adress as adress, people as people,
                //        CASE WHEN ident_dep is not null THEN ident_dep ELSE '' END as ident_dep
                //        from Branchs where id in 
                //        (select distinct b.id from StorageDocs sd left join Branchs b on b.id = sd.id_branch where sd.id_deliv = {0} and b.id_parent = 0
                //        union
                //        select distinct b.id_parent from StorageDocs sd left join Branchs b on b.id = sd.id_branch where sd.id_deliv = {0} and b.id_parent > 0)
                //        order by ident_dep", id_deliv.ToString()), ref ds, null);
                // 
                ExecuteQuery(String.Format("select b.id_parent,b.isolated,b.ident_dep,sd.id,sd.date_doc,sd.id_branch,b.adress as adress,b.people as people,CASE  WHEN b.ident_dep is not null THEN b.ident_dep ELSE '' END as ident_dep from StorageDocs as sd left join Branchs as b on b.id=sd.id_branch where sd.id_deliv={0}", id_deliv.ToString()), ref ds, null);
                //
                //ExecuteQuery("select id,adress,people,ident_dep from Branchs where adress is not null", ref ds, null); // для отладки, нет актуальных данных
                int i;
                /*
                int countIsolated = 0;
                
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if(Convert.ToInt32(ds.Tables[0].Rows[i]["id_parent"])!=0 && Convert.ToInt32(ds.Tables[0].Rows[i]["isolated"])!=0)
                    {
                        ep.SetWorkSheet(1);
                        ep.AddWorkSheet(countIsolated+1);
                        countIsolated++;
                    }
                }
                 
                int numIsolated = 0;
                */ 
                int numSkips = 0;
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                        bool isMainPage = true;
                        /*
                        if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_parent"]) != 0 && Convert.ToInt32(ds.Tables[0].Rows[i]["isolated"]) == 0)
                        {
                            numSkips++;
                            continue;
                        }
                        if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_parent"]) != 0 && Convert.ToInt32(ds.Tables[0].Rows[i]["isolated"]) != 0)
                        {
                            numIsolated++;
                            numSkips++;
                            ep.SetWorkSheet(numIsolated + 1);
                            ep.SetWorkSheetName(numIsolated + 1, ds.Tables[0].Rows[i]["ident_dep"].ToString());
                            isMainPage = false;
                            

                        }
                        else  ep.SetWorkSheet(1);*/
                        if (i == 0 || isMainPage==false)
                        {
                            //
                            ep.SetText(3,1,"Дата: " + String.Format("{0:dd.MM.yyyy}", ds.Tables[0].Rows[i]["date_doc"]));    
                        }
                        int offs = i - numSkips;
                        if (i < ds.Tables[0].Rows.Count - 1 && isMainPage == true)
                        {
                            ep.InsertRow(11 + offs, true);
                        }
                        if (isMainPage == false)
                        {
                            offs = 0;
                        }

                        ep.SetText(11 + offs, 1, (offs + 1).ToString());
                        ep.SetText(11 + offs, 2, " " + ds.Tables[0].Rows[i]["ident_dep"].ToString());
                        String adress = (ds.Tables[0].Rows[i]["adress"] == null) ? "" : ds.Tables[0].Rows[i]["adress"].ToString();
                        //убираем индекс
                        for (int j = 0; j < adress.Length; j++)
                        {
                            if (Char.IsLetter(adress, j) == true)
                            {
                                if (j != 0) adress = adress.Substring(j);
                                break;
                            }
                        }
                        
                        int posTel = adress.IndexOf("тел", 0, StringComparison.CurrentCultureIgnoreCase);
                        if (posTel >= 0) adress = adress.Substring(0, posTel); // убираем телефон
                        String[] a_adr = adress.Split(',');
                        if (a_adr.Length == 4) // область, город, улица, дом
                        {
                            ep.SetText(11 + offs, 8, a_adr[1]);
                            ep.SetText(11 + offs, 9, a_adr[0]);
                            ep.SetText(11 + offs, 12, a_adr[2]);
                            ep.SetText(11 + offs, 13, a_adr[3]);
                        }
                        if (a_adr.Length == 3) // город, улица, дом
                        {
                            ep.SetText(11 + offs, 8, a_adr[0]);
                            ep.SetText(11 + offs, 12, a_adr[1]);
                            ep.SetText(11 + offs, 13, a_adr[2]);
                        }
                        if (a_adr.Length == 2) // город, улица дом (без разделителя)
                        {
                            ep.SetText(11 + offs, 8, a_adr[0]);
                            String[] house = { "д.", "стр.", "c.", "кв." };
                            posTel = -1;
                            for (int j = 0; j < house.Length; j++)
                            {
                                posTel = a_adr[1].IndexOf(house[j], 0, StringComparison.CurrentCultureIgnoreCase);
                                if (posTel >= 0)
                                {
                                    break;
                                }
                            }
                            if (posTel < 0)
                            {
                                ep.SetText(11 + offs, 12, a_adr[1]);
                            }
                            else
                            {
                                ep.SetText(11 + offs, 13, a_adr[1].Substring(posTel));
                                ep.SetText(11 + offs, 12, a_adr[1].Substring(0, posTel));
                            }

                        }
                        
                        String people = (ds.Tables[0].Rows[i]["people"] == null) ? "" : ds.Tables[0].Rows[i]["people"].ToString();
                        posTel = people.IndexOf("тел", 0, StringComparison.CurrentCultureIgnoreCase);
                        String tel = "";
                        String fio = "";
                        if (posTel >= 0)
                        {
                            fio = people.Substring(0, posTel);
                            tel = people.Substring(posTel);
                            for (int j = 0; j < tel.Length; j++)
                            {
                                if (Char.IsDigit(tel, j) == true || tel[j] == '(')
                                {
                                    if (j != 0) tel = tel.Substring(j);
                                    break;
                                }
                            }

                        }
                        else fio = people;
                        ep.SetText(11 + offs, 11, fio);
                        ep.SetText(11 + offs, 14, tel);

                        if (isMainPage == false) ep.SetText(11 + 1, 9, "1");

                    }
                    ep.SetWorkSheet(1);
                    if (i-numSkips >0) ep.SetText(11 + i-numSkips, 9, (i-numSkips).ToString());
                    
                }

            #endregion
            #region книга 124
            if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124
                || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                DateTime dt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);

                object obj = null;
                string fio = sc.UserFIO(User.Identity.Name);

                string fio1 = "";
                int col1 = 6, col2 = 8;

                if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124)
                {
                    col1 = 6; col2 = 8;
                    ExecuteScalar($"select username from storagedocs inner join aspnet_users on StorageDocs.user_id=aspnet_users.id where storagedocs.id={id_doc}", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio = sc.UserFIO((string)obj);
                    ExecuteScalar($"select username from aspnet_users where id=(select id_act from storagedocs where id={id_doc})", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio1 = sc.UserFIO((string)obj);

                }
                if (id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124)
                {
                    col1 = 3; col2 = 5;

                    ExecuteScalar($"select username from storagedocs inner join aspnet_users on StorageDocs.id_act=aspnet_users.id where storagedocs.id={id_doc}", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio = sc.UserFIO((string)obj);
                    ExecuteScalar($"select username from aspnet_users where id=(select user_id from storagedocs where id=(select id_act from storagedocs where id={id_doc}))", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio1 = sc.UserFIO((string)obj);
                }


                ds.Clear();
                SqlCommand comm = new SqlCommand();
                //comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                //res = (string)ExecuteCommand(comm, ref ds, null);
                //ArrayList al = new ArrayList();
                //foreach (DataRow dr in ds.Tables[0].Rows)
                //    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                //ds.Clear();
//                comm.CommandText = "select id_prb, ispin from V_Cards_StorageDocs2 where id_doc=@id";
                comm.CommandText = "select c.isPin, Products.name from Cards_StorageDocs inner join Cards c on Cards_StorageDocs.id_card = c.id inner join Products_Banks p on c.id_prb = p.id inner join Products on p.id_prod = Products.id  where id_doc = @id";
                comm.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                ExecuteCommand(comm, ref ds, null);
                if (ds == null || ds.Tables.Count == 0)
                    return;
                int pin = 0, mc = 0, visa = 0, srv = 0, mir = 0, nfc = 0;
                string company = "";
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, dr["name"].ToString());
                    switch (tp)
                    {
                        case BaseProductType.MasterCard: mc++; break;
                        case BaseProductType.VisaCard: visa++; break;
                        case BaseProductType.ServiceCard: srv++; break;
                        case BaseProductType.MirCard: mir++; break;
                        case BaseProductType.NFCCard: nfc++; break;
                    }
                    if (dr["isPin"] != null && dr["isPin"] != DBNull.Value && Convert.ToBoolean(dr["isPin"]))
                        pin++;
                }
                int i = 0;
                if (mc > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Mastercard");
                    ep.SetText(13 + i, col2, mc.ToString());
                    i++;
                }
                if (visa > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Visa");
                    ep.SetText(13 + i, col2, visa.ToString());
                    i++;
                }
                if (mir > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Mir");
                    ep.SetText(13 + i, col2, mir.ToString());
                    i++;
                }
                if (nfc > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} NFC карты");
                    ep.SetText(13 + i, col2, nfc.ToString());
                    i++;
                }
                if (srv > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Сервисные карты");
                    ep.SetText(13 + i, col2, srv.ToString());
                    i++;
                }
                if (pin > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Пин-конверты");
                    ep.SetText(13 + i, col2, pin.ToString());
                    i++;
                }

                //foreach (MyProd mp in al)
                //{
                //    if (mp.cnts[0] != 0)
                //    {
                //        ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                //        ep.SetText(13 + i, col1, $"{fio1} {mp.Name}");
                //        ep.SetText(13 + i, col2, mp.cnts[0].ToString());
                //        i++;
                //    }
                //}                
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ep.ShowRows(13, 13 + i);

                obj = null;
                ExecuteScalar("select department from Branchs where id=" + sc.BranchId(User.Identity.Name), ref obj, null);
                if (obj != null)
                    ep.SetText(4, 2, obj.ToString());
                ep.SetText(7, 3, fio);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region книга 124 ярлык на мешок
            if (id_type == (int)TypeFormDoc.Book124Label)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                DateTime dt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);
                int id_branch = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]);
                string branch = "";
                int mc = 0, vis = 0, mir = 0, pin = 0;
                using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand comm = conn.CreateCommand())
                    {
                        comm.CommandText = $"select department from Branchs where id={id_branch}";
                        object obj = comm.ExecuteScalar();
                        branch = obj?.ToString();
                        comm.CommandText =
                            $"select prod_name, cnt_perso, id_type from V_Products_StorageDocs where id_doc={id_doc}";
                        SqlDataReader dr = comm.ExecuteReader();
                        
                        while (dr.Read())
                        {
                            cnt = Convert.ToInt32(dr["cnt_perso"]);
                            BaseProductType tp = BranchStore.codeFromTypeAndProdName(Convert.ToInt32(dr["id_type"]), Convert.ToString(dr["prod_name"]));
                            switch (tp)
                            {
                                case BaseProductType.MasterCard:
                                    mc += cnt;
                                    break;
                                case BaseProductType.VisaCard:
                                    vis += cnt;
                                    break;
                                case BaseProductType.MirCard:
                                    mir += cnt;
                                    break;
                                case BaseProductType.PinConvert:
                                    pin += cnt;
                                    break;
                            }
                        }
                        dr.Close();
                    }
                    conn.Close();
                }
                ep.SetText("branch", branch);
                ep.SetText("dt", $"{dt:dd.MM.yyyy}");
                ep.SetText("mc", $"{mc}");
                ep.SetText("mc_r", $"{mc} рублей");
                ep.SetText("vis", $"{vis}");
                ep.SetText("vis_r", $"{vis} рублей");
                ep.SetText("mir", $"{mir}");
                ep.SetText("mir_r", $"{mir} рублей");
                ep.SetText("pin", $"{pin}");
                ep.SetText("pin_r", $"{pin} рублей");
            }

            #endregion
            #region гоз, гоз-подотчет
            if (id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz
                || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                DateTime dt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);

                object obj = null;
                string fio = sc.UserFIO(User.Identity.Name);

                string fio1 = "";
                int col1 = 6, col2 = 8;

                if (id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz)
                {
                    col1 = 6; col2 = 8;
                    ExecuteScalar($"select username from storagedocs inner join aspnet_users on StorageDocs.user_id=aspnet_users.id where storagedocs.id={id_doc}", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio = sc.UserFIO((string)obj);
                    ExecuteScalar($"select username from aspnet_users where id=(select id_act from storagedocs where id={id_doc})", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio1 = sc.UserFIO((string)obj);

                }
                if (id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz)
                {
                    col1 = 3; col2 = 5;

                    ExecuteScalar($"select username from storagedocs inner join aspnet_users on StorageDocs.id_act=aspnet_users.id where storagedocs.id={id_doc}", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio = sc.UserFIO((string)obj);
                    ExecuteScalar($"select username from aspnet_users where id=(select user_id from storagedocs where id=(select id_act from storagedocs where id={id_doc}))", ref obj, null);
                    if (obj != null && obj != DBNull.Value)
                        fio1 = sc.UserFIO((string)obj);
                }


                ds.Clear();
                SqlCommand comm = new SqlCommand();
                //comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                //res = (string)ExecuteCommand(comm, ref ds, null);
                //ArrayList al = new ArrayList();
                //foreach (DataRow dr in ds.Tables[0].Rows)
                //    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                //ds.Clear();
                //                comm.CommandText = "select id_prb, ispin from V_Cards_StorageDocs2 where id_doc=@id";
                comm.CommandText = "select c.isPin, Products.name from Cards_StorageDocs inner join Cards c on Cards_StorageDocs.id_card = c.id inner join Products_Banks p on c.id_prb = p.id inner join Products on p.id_prod = Products.id  where id_doc = @id";
                comm.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                ExecuteCommand(comm, ref ds, null);
                if (ds == null || ds.Tables.Count == 0)
                    return;
                int pin = 0, mc = 0, visa = 0, srv = 0, mir = 0, nfc = 0;
                string company = "";
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, dr["name"].ToString());
                    switch (tp)
                    {
                        case BaseProductType.MasterCard: mc++; break;
                        case BaseProductType.VisaCard: visa++; break;
                        case BaseProductType.ServiceCard: srv++; break;
                        case BaseProductType.MirCard: mir++; break;
                        case BaseProductType.NFCCard: nfc++; break;
                    }
                    if (dr["isPin"] != null && dr["isPin"] != DBNull.Value && Convert.ToBoolean(dr["isPin"]))
                        pin++;
                }
                int i = 0;
                if (mc > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Mastercard");
                    ep.SetText(13 + i, col2, mc.ToString());
                    i++;
                }
                if (visa > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Visa");
                    ep.SetText(13 + i, col2, visa.ToString());
                    i++;
                }
                if (mir > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Mir");
                    ep.SetText(13 + i, col2, mir.ToString());
                    i++;
                }
                if (nfc > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} NFC карты");
                    ep.SetText(13 + i, col2, nfc.ToString());
                    i++;
                }
                if (srv > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Сервисные карты");
                    ep.SetText(13 + i, col2, srv.ToString());
                    i++;
                }
                if (pin > 0)
                {
                    ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                    ep.SetText(13 + i, col1, $"{fio1} Пин-конверты");
                    ep.SetText(13 + i, col2, pin.ToString());
                    i++;
                }

                //foreach (MyProd mp in al)
                //{
                //    if (mp.cnts[0] != 0)
                //    {
                //        ep.SetText(13 + i, 2, $"{dt:dd.MM.yyyy}");
                //        ep.SetText(13 + i, col1, $"{fio1} {mp.Name}");
                //        ep.SetText(13 + i, col2, mp.cnts[0].ToString());
                //        i++;
                //    }
                //}                
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ep.ShowRows(13, 13 + i);

                obj = null;
                ExecuteScalar("select department from Branchs where id=" + sc.BranchId(User.Identity.Name), ref obj, null);
                if (obj != null)
                    ep.SetText(4, 2, obj.ToString());
                ep.SetText(7, 3, fio);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region подотчетные лица
            if (id_type == (int)TypeDoc.SendToPodotchet
                || id_type == (int)TypeDoc.ReceiveToPodotchet
                || id_type == (int)TypeDoc.ReturnFromPodotchet
                || id_type == (int)TypeDoc.ReceiveFromPodotchet
                || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToGozFromPodotchet
                || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToPodotchetFromGoz)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                DateTime dt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);

                ds.Clear();
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                res = (string)ExecuteCommand(comm, ref ds, null);
                ArrayList al = new ArrayList();
                foreach (DataRow dr in ds.Tables[0].Rows)
                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                ds.Clear();
                comm.CommandText = "select id_prb, company from V_Cards_StorageDocs1 where id=@id";
                comm.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                ExecuteCommand(comm, ref ds, null);
                if (ds == null || ds.Tables.Count == 0)
                    return;
                int pin = 0;
                string company = "";
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if (dr["id_prb"] == DBNull.Value)
                        continue;
                    foreach (MyProd mp in al)
                    {
                        if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                        {
                            mp.cnts[0]++; pin++;
                            break;
                        }
                    }
                    //if (dr["company"].ToString().Trim().Length > 1) // чтобы не брать физиков, будем считать что названий компаний из одной буквы нет
                    //    company = dr["company"].ToString().Trim();
                }
                int i = 0;
                foreach (MyProd mp in al)
                {
                    if (mp.cnts[0] != 0)
                    {
                        ep.SetText(10 + i, 1, mp.Name);
                        ep.SetText(10 + i, 6, mp.cnts[0].ToString());
                        i++;
                    }
                }
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ep.SetText(82, 6, pin.ToString());
                ep.ShowRows(10, 10 + i);
                ep.SetText_Name("date_string1", String.Format("{0:dd.MM.yyyy}", dt));
                ep.SetText_Name("date_string2", String.Format("{0:dd.MM.yyyy}", dt));
                if (id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet)
                {
                    ep.SetText_Name("Organization", "");
                    ep.SetText_Name("UPosition", sc.UserPosition(User.Identity.Name));
                    ep.SetText_Name("UFio", String.Format("____________/ {0} /", sc.UserFIO(User.Identity.Name)));
                    ep.SetText_Name("Warrent", "");
                    ep.SetText_Name("WFio", String.Format("____________/________________"));
                }
                if (id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet)
                {
                    ep.SetText_Name("Organization", "");
                    ep.SetText_Name("Warrent", sc.UserPosition(User.Identity.Name));
                    ep.SetText_Name("WFio", String.Format("____________/ {0} /", sc.UserFIO(User.Identity.Name)));
                    ep.SetText_Name("UPosition", "");
                    ep.SetText_Name("UFio", String.Format("____________/________________"));
                }


                object obj = null;
                ExecuteScalar("select department from Branchs where id=" + sc.BranchId(User.Identity.Name), ref obj, null);
                if (obj != null) ep.AddText(90, 1, " (" + obj.ToString() + ")");

                //if (company.Length > 0)
                //{
                //    ds.Clear();
                //    comm.Parameters.Clear();
                //    if (Request.Form["resde"].ToString() != "")
                //        comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where idP=" + Request.Form["resde"].ToString();
                //    else
                //    {
                //        //comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where embosstitle='" + company + "'";
                //        if (branch_main_filial > 0)
                //        {
                //            comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where embosstitle='" + company + "' and BranchMainFilialId=" + branch_main_filial.ToString();
                //        }
                //        else
                //        {
                //            comm.CommandText = "select title, person, position, passport, pdate, pdivision from v_org where embosstitle='" + company + "' and BranchMainFilialId is null";
                //        }
                //    }
                //    ExecuteCommand(comm, ref ds, null);
                //    if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                //    {
                //    }
                //    else
                //    {
                //        ep.SetText_Name("Organization", ds.Tables[0].Rows[0]["title"].ToString().Trim());
                //        ep.SetText_Name("Warrent", String.Format("{0} {5} {1} паспорт {2} от {3:dd.MM.yyyy} выдан {4}", ds.Tables[0].Rows[0]["position"].ToString().Trim(), ds.Tables[0].Rows[0]["person"].ToString().Trim(), ds.Tables[0].Rows[0]["passport"].ToString().Trim(), Convert.ToDateTime(ds.Tables[0].Rows[0]["pdate"]), ds.Tables[0].Rows[0]["pdivision"].ToString().Trim(), ds.Tables[0].Rows[0]["title"].ToString().Trim()));
                //        ep.SetText_Name("WFio", String.Format("____________/{0}", ds.Tables[0].Rows[0]["person"].ToString().Trim()));
                //    }
                //}
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region экспертиза
            if (id_type == (int)TypeDoc.SendToExpertiza || id_type == (int)TypeDoc.ReceiveToFilialExpertiza ||
                id_type == (int)TypeDoc.ReceiveToExpertiza)
            {
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
            }
            #endregion
            #region акт курьерской службу при передаче в филиал
            if (id_type == (int)TypeFormDoc.Courier7777)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                DateTime dt = Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]);

                ds.Clear();
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                res = (string)ExecuteCommand(comm, ref ds, null);
                ArrayList al = new ArrayList();
                foreach (DataRow dr in ds.Tables[0].Rows)
                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                ds.Clear();
                comm.CommandText = "select id_prb, company, ispin from V_Cards_StorageDocs2 where id_doc=@id";
                comm.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                ExecuteCommand(comm, ref ds, null);
                if (ds == null || ds.Tables.Count == 0)
                    return;
                int pin = 0;
                cnt = 0;
                string company = "";
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if (dr["id_prb"] == DBNull.Value)
                        continue;
                    foreach (MyProd mp in al)
                    {
                        if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                        {
                            mp.cnts[0]++;
                            break;
                        }
                    }
                    if (dr["ispin"] != DBNull.Value && Convert.ToBoolean(dr["ispin"]))
                        pin++;
                }
                int i = 0;
                foreach (MyProd mp in al)
                {
                    if (mp.cnts[0] != 0)
                    {
                        ep.SetText(14 + i, 2, $"{i+1}. {mp.Name}");
                        ep.SetText(14 + i, 7, mp.cnts[0].ToString());
                        cnt += mp.cnts[0];
                        i++;
                    }
                }
                ep.SetText(14 + i, 2, "Итого");
                ep.SetText(14 + i, 7, cnt.ToString());
                i++;
                ep.SetWorkSheet(1);
                ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                ep.SetText(42, 7, pin.ToString());
                ep.ShowRows(14, 14 + i);
                ep.SetText_Name("DATE", String.Format("{0:dd.MM.yyyy}", dt));
                ep.SetText_Name("date_string1", String.Format("{0:dd.MM.yyyy}", dt));
                ep.SetText_Name("date_string2", String.Format("{0:dd.MM.yyyy}", dt));

                ep.SetWorkSheet(2);
                ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());

                ep.SetText_Name("date_string3", String.Format("{0:dd.MM.yyyy}", dt));

                DataSet ds1 = new DataSet();
                ds1.Clear();
                //res = ExecuteQuery(String.Format("select fio,pan,company,IdentDepInit,IdentDepCard from V_Cards inner join Cards_StorageDocs on V_Cards.id = Cards_StorageDocs.id_card where id_doc={0} order by IdentDepCard, company, fio", id_sd.ToString()), ref ds1, null);
                res = ExecuteQuery(String.Format("SELECT TOP (100) PERCENT dbo.Cards.id_branchCard, dbo.Cards.fio, dbo.Cards.pan, dbo.Cards.company, dbo.Branchs.ident_dep AS identDepInit, Branchs_1.ident_dep AS identDepCard, CardIdOW FROM dbo.Cards INNER JOIN dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card INNER JOIN dbo.Branchs ON dbo.Cards.id_branchInit = dbo.Branchs.id INNER JOIN dbo.Branchs AS Branchs_1 ON dbo.Cards.id_branchCard = Branchs_1.id WHERE (dbo.Cards_StorageDocs.id_doc = {0}) ORDER BY dbo.Cards.company, dbo.Cards.fio", id_doc.ToString()), ref ds1, null);
                ep.SetFormat(11, 1, 11 + ds1.Tables[0].Rows.Count, 6, "@");
                for (i = 0; i < ds1.Tables[0].Rows.Count; i++)
                {
                    ep.SetText(11 + i, 2, (i + 1).ToString());
                    ep.SetText(11 + i, 3, ds1.Tables[0].Rows[i]["fio"].ToString().Trim());
                    ep.SetText(11 + i, 4, ds1.Tables[0].Rows[i]["CardIdOW"].ToString().Trim());
                    ep.SetText(11 + i, 5, ds1.Tables[0].Rows[i]["pan"].ToString().Trim());
                    ep.SetText(11 + i, 6, ds1.Tables[0].Rows[i]["company"].ToString().Trim());
                    ep.SetText(11 + i, 7, ds1.Tables[0].Rows[i]["IdentDepInit"].ToString().Trim());
                    ep.SetText(11 + i, 8, ds1.Tables[0].Rows[i]["IdentDepCard"].ToString().Trim());
                }
                //            ep.SetRangeAutoFit(8, 1, 8 + ds1.Tables[0].Rows.Count, 6); // чтобы влазило на страницу
                ep.SetRangeBorders(11, 2, 11 + ds1.Tables[0].Rows.Count, 8);
            }
            #endregion
        }
        private void ZapExcelUz(int id_type, ExcelAp ep)
        {
            string dt_doc = "";
            int col = 0, row = 0, sum = 0, cnt = 0;
            DataSet ds = new DataSet();
            int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
            dt_doc = String.Format("{0:dd.MM.yyyy}", Convert.ToDateTime(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["date_doc"]));
            #region Персонализация
            if (id_type == (int)TypeDoc.PersoCard)
            {
                ep.SetText_Name("dt_doc", dt_doc);
                ArrayList banks = new ArrayList();
                ArrayList prods = new ArrayList();
                ds.Clear();
                res = ExecuteQuery($"select distinct bank_name from V_Products_StorageDocs where id_doc={id_doc}", ref ds, null);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    banks.Add(ds.Tables[0].Rows[i]["bank_name"].ToString());
                ds.Clear();
                res = ExecuteQuery($"select distinct prod_name from V_Products_StorageDocs where id_doc={id_doc} and id_type=1", ref ds, null);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    prods.Add(ds.Tables[0].Rows[i]["prod_name"].ToString());                
                foreach (string b in banks)
                {
                    ep.SetText(16 + row, 1, b);
                    col = 2;
                    sum = 0;
                    foreach (string p in prods)
                    {
                        ep.SetText(15, col, p);
                        for (int j = 0; j < gvProducts.Rows.Count; j++)
                        {
                            if (gvProducts.DataKeys[j].Values["prod_name"].ToString() == p && gvProducts.DataKeys[j].Values["bank_name"].ToString() == b)
                            {
                                try
                                {
                                    cnt = Convert.ToInt32(gvProducts.DataKeys[j].Values["cnt_perso"]);
                                }
                                catch
                                {
                                    cnt = 0;
                                }
                                if (cnt > 0)
                                    ep.SetText(16 + row, col, cnt.ToString());
                                sum += cnt;
                            }
                            
                        }
                        col++;
                    }
                    ep.SetText(15, col, "Итого");
                    ep.SetText(16 + row, col, sum.ToString());

                    row++;
                }
                ep.SetRangeBorders(15, 1, 16 + row + 1, col);
                ep.SetRangeBorders(123, 1, 123, col);
                for (int j = 0; j < 50; j++)
                    ep.SetText(123, col + 1 + j, "");
                ep.ShowRows(16, 16 + row);
            }
            #endregion
            #region Персонализация, записка на выпуск карт (там где много листов)
            if (id_type == (int)TypeFormDoc.OfficeNote)
            {
                ds.Clear();
                res = ExecuteQuery($"select prod_name, cnt_perso, cnt_brak, bank_name, id_type from V_Products_StorageDocs where id_doc={id_doc} order by bank_name, prod_name", ref ds, null);

                ep.SetWorkSheet(1);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"].ToString()) != 1)
                        continue;
                    ep.SetText(13 + row, 1, ds.Tables[0].Rows[i]["bank_name"].ToString());
                    ep.SetText(13 + row, 3, ds.Tables[0].Rows[i]["prod_name"].ToString());
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                    }
                    catch
                    {
                        cnt = 0;
                    }
                    ep.SetText(13 + row, 6, (cnt > 0) ? cnt.ToString() : "");
                    sum += cnt;
                    row++;
                }
                ep.SetText(94, 6, sum.ToString());
                ep.SetText_Name("date1_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date1_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date1_2", $"{dt_doc:dd.MM.yyyy}");
                ep.ShowRows(13, 13 + row);               
                ep.SetWorkSheet(2);
                sum = 0;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"].ToString()) != 2)
                        continue;
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                    }
                    catch { cnt = 0; }
                    sum += cnt;
                }
                ep.SetText(22, 8, sum.ToString());
                ep.SetText_Name("date2_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date2_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date2_2", $"{dt_doc:dd.MM.yyyy}");
                ep.SetWorkSheet(3);
                row = 0;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"].ToString()) != 1)
                        continue;
                    ep.SetText(13 + row, 1, ds.Tables[0].Rows[i]["bank_name"].ToString());
                    ep.SetText(13 + row, 3, ds.Tables[0].Rows[i]["prod_name"].ToString());
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                    }
                    catch
                    {
                        cnt = 0;
                    }
                    ep.SetText(13 + row, 6, (cnt > 0) ? cnt.ToString() : "");
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"].ToString());
                    }
                    catch
                    {
                        cnt = 0;
                    }
                    ep.SetText(13 + row, 9, (cnt > 0) ? cnt.ToString() : "");
                    row++;
                }                
                ep.SetText_Name("date3_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date3_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date3_2", $"{dt_doc:dd.MM.yyyy}");
                ep.ShowRows(13, 13 + row);
                ep.SetWorkSheet(4);
                sum = 0;
                int sum1 = 0;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"].ToString()) != 2)
                        continue;
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                    }
                    catch { cnt = 0; }
                    sum += cnt;
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"].ToString());
                    }
                    catch { cnt = 0; }
                    sum1 += cnt;
                }
                ep.SetText(22, 8, sum.ToString());
                ep.SetText(22, 9, sum1.ToString());
                ep.SetText_Name("date4_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date4_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date4_2", $"{dt_doc:dd.MM.yyyy}");
                ep.SetWorkSheet(5);
                row = 0; sum = 0;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"].ToString()) != 1)
                        continue;
                    ep.SetText(13 + row, 1, ds.Tables[0].Rows[i]["bank_name"].ToString());
                    ep.SetText(13 + row, 3, ds.Tables[0].Rows[i]["prod_name"].ToString());
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                    }
                    catch
                    {
                        cnt = 0;
                    }
                    ep.SetText(13 + row, 6, (cnt > 0) ? cnt.ToString() : "");
                    sum += cnt;
                    row++;
                }
                ep.SetText(94, 6, sum.ToString());
                ep.SetText_Name("date5_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date5_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date5_2", $"{dt_doc:dd.MM.yyyy}");
                ep.ShowRows(13, 13 + row);
                ep.SetWorkSheet(6);
                sum = 0;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"].ToString()) != 2)
                        continue;
                    try
                    {
                        cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                    }
                    catch { cnt = 0; }
                    sum += cnt;
                }
                ep.SetText(22, 8, sum.ToString());
                ep.SetText_Name("date6_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date6_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date6_2", $"{dt_doc:dd.MM.yyyy}");
                ep.SetWorkSheet(7);
                ds.Clear();
                res = ExecuteQuery($"select dbo.BranchBank(id_BranchCard) as bank, dbo.BranchMainDep(id_BranchCard) as mfo, dbo.BranchDep(id_BranchCard) as agent, dbo.Productname(id_prb) as product, count(*) as perso from V_Cards_StorageDocs where id_doc={id_doc} group by id_BranchCard, id_prb order by bank, mfo, agent, product", ref ds, null);
                row = 0;
                for (int i=0;i<ds.Tables[0].Rows.Count;i++)
                {
                    ep.SetText(8+row, 1, $"{i + 1}");
                    ep.SetText(8+row, 2, $"{ds.Tables[0].Rows[i]["bank"]}");
                    ep.SetText(8+row, 3, $"{ds.Tables[0].Rows[i]["mfo"]}");
                    ep.SetText(8+row, 4, $"{ds.Tables[0].Rows[i]["agent"]}");
                    ep.SetText(8+row, 5, $"{ds.Tables[0].Rows[i]["product"]}");
                    ep.SetText(8+row, 6, $"{ds.Tables[0].Rows[i]["perso"]}");
                    ep.SetText(8+row, 7, $"{ds.Tables[0].Rows[i]["perso"]}");
                    row++;
                }
                ep.ShowRows(8, 8 + row);
                ep.SetText_Name("date7_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date7_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date7_2", $"{dt_doc:dd.MM.yyyy}");
                ep.SetWorkSheet(8);
                ds.Clear();
                res = ExecuteQuery($"select dbo.BranchBank(id_BranchCard) as bank, dbo.BranchMainDep(id_BranchCard) as mfo, dbo.BranchDep(id_BranchCard) as agent, dbo.Productname(id_prb) as product, count(*) as perso from V_Cards_StorageDocs where id_doc={id_doc} group by id_BranchCard, id_prb order by bank, mfo, agent, product", ref ds, null);
                row = 0;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ep.SetText(8 + row, 1, $"{i + 1}");
                    ep.SetText(8 + row, 2, $"{ds.Tables[0].Rows[i]["bank"]}");
                    ep.SetText(8 + row, 3, $"{ds.Tables[0].Rows[i]["mfo"]}");
                    ep.SetText(8 + row, 4, $"{ds.Tables[0].Rows[i]["agent"]}");
                    ep.SetText(8 + row, 5, $"{ds.Tables[0].Rows[i]["product"]}");
                    ep.SetText(8 + row, 6, $"{ds.Tables[0].Rows[i]["perso"]}");
                    ep.SetText(8 + row, 7, $"{ds.Tables[0].Rows[i]["perso"]}");
                    row++;
                }
                ep.ShowRows(8, 8 + row);
                ep.SetText_Name("date8_str", $"Дата {dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date8_1", $"{dt_doc:dd.MM.yyyy}");
                ep.SetText_Name("date8_2", $"{dt_doc:dd.MM.yyyy}");

                ds.Clear();
                ExecuteQuery($"select distinct dbo.BranchBank(id_branchcard) as b from V_Cards_StorageDocs where id_doc={id_doc}", ref ds, null);
                for (int i = 0; i < ds.Tables[0].Rows.Count - 1; i++)
                {
                    ep.SetWorkSheet(9);
                    ep.AddWorkSheet(9);
                }
                ArrayList banks = new ArrayList();
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    banks.Add(ds.Tables[0].Rows[i]["b"].ToString());
                }
                int index = 0;
                foreach(string b in banks)
                {
                    ep.SetWorkSheet(9 + index);
                    ds.Clear();
                    row = 0;
                    ExecuteQuery($"select dbo.BranchMainDep(id_BranchCard) as mfo, dbo.BranchDep(id_BranchCard) as agent, dbo.Productname(id_prb) as product, pan, fio from V_Cards_StorageDocs where id_doc={id_doc} and dbo.BranchBank(id_branchCard)='{b}' order by mfo, agent, product, pan_unmasked, pan", ref ds, null);
                    ep.SetText(5, 5, b);
                    for (int i=0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        ep.SetText(8 + row, 1, $"{i+1}");
                        ep.SetText(8 + row, 2, ds.Tables[0].Rows[i]["mfo"].ToString());
                        ep.SetText(8 + row, 3, ds.Tables[0].Rows[i]["agent"].ToString());
                        ep.SetText(8 + row, 4, ds.Tables[0].Rows[i]["product"].ToString());
                        ep.SetText(8 + row, 5, ds.Tables[0].Rows[i]["pan"].ToString());
                        ep.SetText(8 + row, 9, ds.Tables[0].Rows[i]["fio"].ToString());
                        row++;
                    }
                    ep.SetRangeBorders(8, 1, 8 + row, 10);
                    index++;
                }
                ep.SetWorkSheet(9 + index);
                ds.Clear();
                row = 0;
                ExecuteQuery($"select dbo.BranchBank(id_BranchCard) as bank, dbo.BranchMainDep(id_BranchCard) as mfo, dbo.BranchDep(id_BranchCard) as agent, dbo.Productname(id_prb) as product, pan, fio, passport from V_Cards_StorageDocs where id_doc={id_doc} order by pan_unmasked, pan, agent, bank", ref ds, null);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {                    
                    ep.SetText(2 + row, 1, ds.Tables[0].Rows[i]["pan"].ToString());
                    ep.SetText(2 + row, 2, ds.Tables[0].Rows[i]["fio"].ToString());
                    ep.SetText(2 + row, 3, ds.Tables[0].Rows[i]["bank"].ToString());
                    ep.SetText(2 + row, 4, ds.Tables[0].Rows[i]["mfo"].ToString());
                    ep.SetText(2 + row, 5, ds.Tables[0].Rows[i]["agent"].ToString());
                    ep.SetText(2 + row, 6, $"{dt_doc:dd.MM.yyyy}");
                    string pass = (ds.Tables[0].Rows[i]["passport"] == null || ds.Tables[0].Rows[i]["passport"] == DBNull.Value) ? "" : ds.Tables[0].Rows[i]["passport"].ToString();
                    ep.SetText(2 + row, 7, pass.Split(' ')?[0]);
                    ep.SetText(2 + row, 8, (pass.Split(' ' ).Length > 1) ? pass.Split(' ')[1] : "");
                    ep.SetText(2 + row, 9, ds.Tables[0].Rows[i]["product"].ToString());
                    ep.SetText(2 + row, 10, ds.Tables[0].Rows[i]["pan"].ToString().Substring(0, 9).Replace(" ",""));                    
                    row++;
                }

            }
            #endregion
            #region Приложение6 - SendToFilial
            if (id_type == (int)TypeDoc.SendToFilial)
            {
                int id_deliv = 0;
                if (gvDocs.Rows.Count > 0)
                    id_deliv = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_deliv"]);
                if (id_deliv == 0) // не рассылка
                {
                    ep.SetWorkSheet(1);
                    ep.SetWorkSheetName(1, "a_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                    ExcelDelivFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]), gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString(), ep);
                    ep.SetWorkSheet(2);
                    ep.SetWorkSheetName(2, "r_" + gvDocs.DataKeys[gvDocs.SelectedIndex].Values["branch"].ToString());
                    ExcelReestrFilial(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), ep);
                }
                if (id_deliv > 0) //рассылка
                {
                    ds.Clear();
                    res = ExecuteQuery(String.Format("select id,id_branch,branch from V_StorageDocs where id_deliv={0} order by branch", id_deliv.ToString()), ref ds, null);
                    int countBranchs = 0;


                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        int[] branchs = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), false);
                        //int[] branchsIsolated = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), true);
                        if (branchs.Length > 0) countBranchs++;
                        //if (branchsIsolated.Length > 0) countBranchs += branchsIsolated.Length;
                    }

                    for (int i = 0; i < countBranchs; i++)
                    {
                        ep.SetWorkSheet(1);
                        if (i != countBranchs - 1) ep.AddWorkSheet();
                        ep.SetWorkSheet(2);
                        if (i != countBranchs - 1) ep.AddWorkSheet();
                    }

                    for (int i = 0, n = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        int[] branchs = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), false);
                        //BranchStore[] branchsIsolated = getCardBranchFromStoreDocs(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]));
                        if (branchs.Length > 0)
                        {
                            ep.SetWorkSheet(2 * n + 1);
                            ep.SetWorkSheetName(2 * n + 1, "a_" + ds.Tables[0].Rows[i]["branch"].ToString());
                            ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]), ds.Tables[0].Rows[i]["branch"].ToString(), ep);
                            ep.SetWorkSheet(2 * n + 2);
                            ep.SetWorkSheetName(2 * n + 2, "r_" + ds.Tables[0].Rows[i]["branch"].ToString());
                            ExcelReestrFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), ep);
                            n++;
                        }
                        //if (branchsIsolated.Length > 0)
                        //{
                        //    for (int j = 0; j < branchsIsolated.Length; j++)
                        //    {
                        //        ep.SetWorkSheet(2 * n + 1);
                        //        ep.SetWorkSheetName(2 * n + 1, "a_" + branchsIsolated[j].department);
                        //        ExcelDelivFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), branchsIsolated[j].id, branchsIsolated[j].department, ep, null, branchsIsolated[j]);
                        //        ep.SetWorkSheet(2 * n + 2);
                        //        ep.SetWorkSheetName(2 * n + 2, "r_" + branchsIsolated[j].department);
                        //        ExcelReestrFilial(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), ep, null, branchsIsolated[j]);
                        //        n++;
                        //    }
                        //}
                    }
                    //if (id_type == (int)TypeFormDoc.Akt2Reestr)
                    //{
                    //    for (int i = 0; i < countBranchs; i++)
                    //    {
                    //        ep.SetWorkSheet(2 * i + 1 + i);
                    //        ep.AddWorkSheet(2 * i + 1 + i);
                    //    }
                    //}

                }
            }
            #endregion
        }
        private int[] getCardBranchFromStoreDocs(int id_doc, bool isolated)
        {
            DataSet ds = new DataSet();
            ExecuteQuery("select distinct c.id_BranchCard as id from Cards_StorageDocs cs " +
                "left join Cards c on cs.id_card=c.id " +
                "left join Branchs b on c.id_BranchCard=b.id " +
                "where id_doc=" + id_doc.ToString() + " and b.isolated" + ((isolated==true) ? ">0":"<1"), ref ds, null);
            int[] idBranches;
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0) idBranches = new int[ds.Tables[0].Rows.Count];
            else idBranches = new int[0];
            for (int i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
            {
                idBranches[i]=Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);
            }
            return idBranches;
        }

        private BranchStore[] getCardBranchFromStoreDocs(int id_doc)
        {
            DataSet ds = new DataSet();
            ExecuteQuery("select distinct c.id_BranchCard as id, b.ident_dep, b.department from Cards_StorageDocs cs " +
                "left join Cards c on cs.id_card=c.id " +
                "left join Branchs b on c.id_BranchCard=b.id " +
                "where id_doc=" + id_doc.ToString() + " and b.isolated>0", ref ds, null);
            BranchStore[] Branches;
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0) Branches = new BranchStore[ds.Tables[0].Rows.Count];
            else Branches = new BranchStore[0];
            for (int i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
            {
                Branches[i]=new BranchStore(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]),ds.Tables[0].Rows[i]["ident_dep"].ToString(),ds.Tables[0].Rows[i]["department"].ToString());
            }
            return Branches;
        }


        protected void lbProduct_Click(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbViewP.Text = "P";
                ViewTypePanel();
            }
        }
        protected void lbCard_Click(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbViewP.Text = "C";
                ViewTypePanel();
            }
        }
        protected void gvProducts_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                SetButtonProd();
                lbInform.Text = "";
                gvProducts.Rows[gvProducts.SelectedIndex].Focus();
            }
        }
        protected void bAutoD_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (!User.Identity.IsAuthenticated)
                    Response.Redirect("~\\Account\\Unauthenticated.aspx", true);

                AddStorageDoc();
            }
        }
        protected void bEditD_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (!User.Identity.IsAuthenticated)
                    Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
                Refr(gvDocs.SelectedIndex, false);
                lbInform.Text = "";
            }
        }
        protected void bDeleteD_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);

            lock (Database.lockObjectDB)
            {


                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                int id_deliv = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_deliv"]);
                string nom = gvDocs.DataKeys[gvDocs.SelectedIndex].Values["number_doc"].ToString();
                int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);

                SqlCommand comm = conn.CreateCommand();
                comm.CommandText = "select priz_gen from StorageDocs where id=" + id_doc.ToString();
                object obj = comm.ExecuteScalar();
                if (Convert.ToBoolean(obj))
                {
                    lbInform.Text = "Нельзя удалить подтвержденный документ";
                    return;
                }

                if (id_deliv == 0)
                {
                    if (id_type == (int)TypeDoc.SendToBank) //~! && isMainFilial() == false)
                    {
                        comm.Parameters.Clear();
                        comm.CommandText = "select count(*) from V_Cards_StorageDocs where id_doc=@doc and id_prop!=1"; //~!+ ((isMainFilial() == false) ? "" : " and id_prop!=10");
                        comm.Parameters.Add("@doc", SqlDbType.Int).Value = id_doc;
                        int cnt = Convert.ToInt32(comm.ExecuteScalar());
                        if (Convert.ToInt32(cnt) > 0)
                        {
                            lbInform.Text = "В документе есть карта без состояния ОК";
                            return;
                        }
                    }
                    //~!if (id_type == (int)TypeDoc.SendToBank && isMainFilial() == true)
                    //~!{
                        /*
                         comm.Parameters.Clear();
                         comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                         comm.Parameters.Add("@dt", SqlDbType.DateTime);
                         comm.Parameters["@dt"].Value = DBNull.Value;
                         comm.CommandText = "Update Cards set id_stat=5, dateGetTerminate=@dt,id_BranchCurrent=id_BranchCard where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop <> 1";
                         comm.ExecuteNonQuery();
                        */
                    //~!}
                    WebLog.LogClass.WriteToLog("StorDoc.bDeleteD_Click Start id_doc={0}, id_type={1}, user = {2}, userbranch = {3}", id_doc, id_type, User.Identity.Name, sc.BranchId(User.Identity.Name));
                    DelDoc(id_doc);
                    WebLog.LogClass.WriteToLog("StorDoc.bDeleteD_Click End id_doc={0}, id_type={1}, user = {2}, userbranch = {3}", id_doc, id_type, User.Identity.Name, sc.BranchId(User.Identity.Name));
                    WebLog.LogClass.WriteToLog(sc.UserGuid(User.Identity.Name), String.Format("Удален документ {0}", nom), null);

                }
                else
                {
                    DataSet ds1 = new DataSet();
                    ds1.Clear();
                    res = ExecuteQuery(String.Format("select id,type,priz_gen from StorageDocs where id_deliv={0}", id_deliv.ToString()), ref ds1, null);

                    for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                    {
                        id_doc = Convert.ToInt32(ds1.Tables[0].Rows[i]["id"]);
                        WebLog.LogClass.WriteToLog(sc.BranchId(User.Identity.Name) + " " + User.Identity.Name + " " + sc.UserGuid(User.Identity.Name) + "\t" + "Удаление документа (цикл) " + id_doc.ToString() + " " + nom + " " + id_type);
                        DelDoc(id_doc);
                        WebLog.LogClass.WriteToLog(sc.UserGuid(User.Identity.Name), String.Format("Удален документ {0}", nom), null);
                    }
                }
                Refr(0, false);
            }
        }
        private void DelDoc(int id_doc)
        {
            #region реестр курьерской службы
            try {                
                int id_type = 0, number_doc = 0;
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand comm = conn.CreateCommand())
                    {
                        comm.CommandText = "select number_doc, type from StorageDocs where id=" + id_doc.ToString();
                        using (SqlDataReader dr = comm.ExecuteReader())
                        { 
                            if (dr.Read())
                            {
                                id_type = Convert.ToInt32(dr["type"]);
                                number_doc = Convert.ToInt32(dr["number_doc"]);
                            }
                            dr.Close();
                        }
                    }
                    conn.Close();
                }
                if (id_type == 8 || id_type == 5)
                {
                    documentJson.Clear();
                    string[] branchs = (ConfigurationManager.AppSettings.AllKeys.Contains("CourierBranches")) ? ConfigurationManager.AppSettings["CourierBranches"].Split(',') : new string[] { };
                    using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = $"select Cards.pan, cards.CardIdOW, Branchs.ident_dep from Cards_StorageDocs left join Cards on Cards_StorageDocs.id_card=Cards.id left join Branchs on Cards.id_BranchCard=Branchs.id where Cards_StorageDocs.id_doc={id_doc}";
                            using (SqlDataReader dr = comm.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    CardData c = new CardData();
                                    c.externalCardId = (dr["CardIdOW"] == DBNull.Value) ? "" : dr["CardIdOW"].ToString();
                                    c.number = (dr["pan"] == DBNull.Value) ? "" : dr["pan"].ToString().Trim();
                                    c.branch = (dr["ident_dep"] == DBNull.Value) ? "" : dr["ident_dep"].ToString();
                                    if (branchs.Contains(c.branch))
                                        documentJson.cards.Add(c);
                                }
                                dr.Close();
                            }
                        }
                        conn.Close();
                    }
                    if (documentJson.cards.Count > 0)
                    {
                        documentJson.documentId = number_doc.ToString();
                        documentJson.operationType = "Delete";
                        documentJson.cardsState = (id_type == 8) ? "Обработка" : "Хранилище";
                        DateTime dtjson = DateTime.Now;
                        documentJson.dateTime = dtjson.ToString("yyyy-MM-ddTHH:mmK");
                        var serializer = new Newtonsoft.Json.JsonSerializer();
                        var stringWriter = new StringWriter();
                        using (var writer = new Newtonsoft.Json.JsonTextWriter(stringWriter))
                        {
                            writer.QuoteName = false;
                            writer.Formatting = Newtonsoft.Json.Formatting.Indented;
                            serializer.Serialize(writer, documentJson);
                        }
                        var filename = Path.Combine(ConfigurationManager.AppSettings["CourierFolder"],
                            $"{number_doc}_{dtjson.ToString("yyyyMMddHHmm")}.json");
                        using (StreamWriter sw = new StreamWriter(filename))
                        {
                            sw.WriteLine(stringWriter.ToString());
                            sw.Close();
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                WebLog.LogClass.WriteToLog(exp.ToString());
                lbInform.Text = "Ошибка создания реестра курьерской службы";
                return;
            }
            #endregion
            try
            {
                SqlCommand sqCom = new SqlCommand();

                sqCom.CommandText = "delete from Products_StorageDocs where id_doc=@id";
                WebLog.LogClass.WriteToLog(sc.BranchId(User.Identity.Name) + " " + User.Identity.Name + " " + sc.UserGuid(User.Identity.Name) + "\t" + sqCom.CommandText + "\t" + id_doc.ToString()); 
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                res = ExecuteNonQuery(sqCom, null);

                sqCom.CommandText = "delete from Cards_StorageDocs where id_doc=@id";
                WebLog.LogClass.WriteToLog(sc.BranchId(User.Identity.Name) + " " + User.Identity.Name + " " +
                                    sc.UserGuid(User.Identity.Name) + "\t" + sqCom.CommandText + "\t" +
                                    id_doc.ToString()); 
                sqCom.Parameters["@id"].Value = id_doc;
                res = ExecuteNonQuery(sqCom, null);

                sqCom.CommandText = "delete from AccountablePerson_StorageDocs where id_doc=@id";
                WebLog.LogClass.WriteToLog(sc.BranchId(User.Identity.Name) + " " + User.Identity.Name + " " +
                                    sc.UserGuid(User.Identity.Name) + "\t" + sqCom.CommandText + "\t" +
                                    id_doc.ToString());
                sqCom.Parameters["@id"].Value = id_doc;
                res = ExecuteNonQuery(sqCom, null);
                sqCom.CommandText = "delete from StorageDocs where id=@id";
                res = ExecuteNonQuery(sqCom, null);
            }
            catch (Exception ex)
            {
                WebLog.LogClass.WriteToLog(ex.ToString());
            }
        }
        protected void bConfFieldD_Click(object sender, ImageClickEventArgs e)
        {
            Response.Redirect("~/StorDoc.aspx");
        }
        protected void bDelProd_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvProducts.DataKeys[gvProducts.SelectedIndex].Values["id"]);
                int prb_id = Convert.ToInt32(gvProducts.DataKeys[gvProducts.SelectedIndex].Values["id_prb"]);
                int doc_id = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);

                SqlCommand sqCom = conn.CreateCommand();

                sqCom.CommandText = "delete from Products_StorageDocs where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                ExecuteNonQuery(sqCom, null);

                //возврат с упаковки, персонализация, отправка на упаковку
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == 17
                    || Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == 8
                    || Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == 16)
                {
                    sqCom.Parameters.Clear();
                    sqCom.CommandText = "delete from Cards_StorageDocs where id_doc=@did and id_card in (select id_card from V_Cards_StorageDocs where id_doc=@did and id_prb=@pid)";
                    sqCom.Parameters.Add("@did", SqlDbType.Int).Value = doc_id;
                    sqCom.Parameters.Add("@pid", SqlDbType.Int).Value = prb_id;
                    ExecuteNonQuery(sqCom, null);
                    AddAutoStorageProduct(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]), null);
                    ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                }

                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                lbInform.Text = "";
            }
        }
        protected void bNewProd_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                lbInform.Text = "";
            }
        }
        protected void bEditProd_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvProducts.SelectedIndex);
                lbInform.Text = "";
            }
        }
        protected void gvCards_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvCards.Rows[gvCards.SelectedIndex].Focus();
                if (gvCards.DataKeys[gvCards.SelectedIndex].Values["id"] == DBNull.Value)
                {                   
                    allcards = true;
                    ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                }
                SetButtonCard();
                lbInform.Text = "";
            }
        }
        protected void bDelCard_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
                    Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvCards.DataKeys[gvCards.SelectedIndex].Values["id"]);
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                SqlCommand sqCom = conn.CreateCommand();

                // Проверка на подтвержденность
                sqCom.CommandText = "select top 1 priz_gen from StorageDocs where id=@id_doc";
                sqCom.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                Object priz_gen = sqCom.ExecuteScalar();
                if(priz_gen!=DBNull.Value && Convert.ToBoolean(priz_gen)==true)
                {
                    WebLog.LogClass.WriteToLog("StorDoc.bDelCard_Click - нельзя удалить карту из подтвержденного документа");
                    ClientScript.RegisterClientScriptBlock(GetType(), "err_priz_gen", "<script type='text/javascript'>$(document).ready(function(){ ShowError('Нельзя удалить карту из подтвержденного документа');});</script>");
                    return;
                }


                WebLog.LogClass.WriteToLog("StorDoc.bDelCard_Click Start id_doc={0}, id_card={1}, user = {2}, userbranch = {3}", id_doc, id, User.Identity.Name, sc.BranchId(User.Identity.Name));
                sqCom.CommandText = "delete from Cards_StorageDocs where id=@id";
                
                sqCom.Parameters.Clear();

                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                //res = ExecuteNonQuery(sqCom, null);
                sqCom.ExecuteNonQuery();
                WebLog.LogClass.WriteToLog("StorDoc.bDelCard_Click End id_doc={0}, id_card={1}, user = {2}, userbranch = {3}", id_doc, id, User.Identity.Name, sc.BranchId(User.Identity.Name));
                AddAutoStorageProduct(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]), null);
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == 17)
                {
                    //AddAutoStorageProduct(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 17, null);
                    ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                }
                

                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                lbInform.Text = "";
            }
        }
        protected void bNewCard_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            lock (Database.lockObjectDB)
            {
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == 13) //filial-filial
                {
                    int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    AddAutoStorageProduct(id_doc, id_type, null);
                    ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                    lbInform.Text = "";
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.SendToBank) //возврат банка в банк
                {
                    int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    AddAutoStorageProduct(id_doc, id_type, null);
                    ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                    lbInform.Text = "";  
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.SendToPodotchet)
                {
                    //int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    //int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    //AddAutoStorageProduct(id_doc, id_type, null);
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ToBook124 ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.FromBook124 ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.GetBook124 ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ReceiveBook124)
                {
                    int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    AddAutoStorageProduct(id_doc, id_type, null);
                    ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                    lbInform.Text = "";
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ToGoz ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.FromGoz ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.GetGoz ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ReceiveGoz)
                {
                    int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    AddAutoStorageProduct(id_doc, id_type, null);
                    ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                    lbInform.Text = "";
                }
                if (Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ToGozFromPodotchet ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.FromGozToPodotchet ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.FromPodotchetToGoz ||
                    Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]) == (int)TypeDoc.ToPodotchetFromGoz)
                {
                    int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                    int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                    AddAutoStorageProduct(id_doc, id_type, null);
                    ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                    lbInform.Text = "";
                }


                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                SetButtonDoc();
                SetButtonCard();
                lbInform.Text = "";
            }
        }
        protected void bDeleteD2_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            lock (Database.lockObjectDB)
            {
                if (!User.Identity.IsAuthenticated)
                {
                    Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
                }
                int id = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                SqlCommand sqCom = conn.CreateCommand();
                sqCom.CommandText = "delete from Products_StorageDocs where id_doc=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                //res = ExecuteNonQuery(sqCom, null);
                sqCom.ExecuteNonQuery();
                sqCom.CommandText = "delete from Cards_StorageDocs where id_doc=@id";
                sqCom.Parameters["@id"].Value = id;
                //            res = ExecuteNonQuery(sqCom, null);
                sqCom.ExecuteNonQuery();
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                lbInform.Text = "";
            }
        }
        protected void bAutoProd_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            lock (Database.lockObjectDB)
            {
                int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                int id_type = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
                AddAutoStorageProduct(id_doc, id_type, null);
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                lbInform.Text = "";
            }
        }
        protected void bEditCard_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
                Response.Redirect("~\\Account\\Unauthenticated.aspx", true);
            lock (Database.lockObjectDB)
            {
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.SelectedIndex);
                lbInform.Text = "";
            }
        }
        private void SendConfirm(Restrictions r, int id_doc)
        {
            docTempName = "";
            if (r == Restrictions.ConfirmPerso)
            {
                WorkExcel((int)TypeFormDoc.OfficeNote, false);
                if (docTempName.Length > 0)
                {
                    File.Copy(docTempName, String.Format("{0}SZ_{1:ddMMyyyy}.xls", ConfigurationSettings.AppSettings["ArchivePath"].ToString(), DateTime.Now), true);
                    MailMessage mm = new MailMessage();
                    DataSet ds1 = new DataSet();
                    res = ExecuteQuery(String.Format("select UserName from V_UserAction where ActionID={0}", (int)Restrictions.ConfirmPerso), ref ds1, null);
                    if (ds1.Tables.Count == 0)
                        return;
                    for (int t = 0; t < ds1.Tables[0].Rows.Count; t++)
                    {
                        object obj = null;
                        ExecuteScalar(String.Format("select EMail from vw_aspnet_MembershipUsers where UserName='{0}'", ds1.Tables[0].Rows[t]["UserName"].ToString()), ref obj, null);
                        if (obj != null)
                        {
                            try
                            {
                                MailAddress mTo = new MailAddress((string)obj);
                                mm.Bcc.Add(mTo);
                            }
                            catch { }
                        }
                    }
                    try
                    {
                        SmtpClient sc = new SmtpClient(ConfigurationSettings.AppSettings["SmtpServer"]);
                        sc.Credentials = new NetworkCredential(ConfigurationSettings.AppSettings["EMailFrom"], ConfigurationSettings.AppSettings["Pwd"]);
                        MailAddress mailFrom = new MailAddress(ConfigurationSettings.AppSettings["EMailFrom"], "CardPerso");
                        mm.From = mailFrom;
                        mm.Subject = "CardPerso - служебная записка на выпуск";
                        mm.Body = String.Format("В системе CardPerso сформирована служебная записка на выпуск.\n{0:HH.mm dd.MM.yyyy}", DateTime.Now);
                        mm.Attachments.Add(new Attachment(docTempName));
                        if (mm.Bcc.Count > 0)
                        {
                            sc.Send(mm);
                            lbInform.Text = "Служебная записка отправлена по рассылке";
                        }
                    }
                    catch { }
                }
            }
            if (r == Restrictions.InformProduction)
            {
                MailMessage mm = new MailMessage();
                DataSet ds1 = new DataSet();
                res = ExecuteQuery(String.Format("select UserName from V_UserAction where ActionID={0}", (int)Restrictions.InformProduction), ref ds1, null);
                if (ds1.Tables.Count == 0)
                    return;
                for (int t = 0; t < ds1.Tables[0].Rows.Count; t++)
                {
                    object obj = null;
                    ExecuteScalar(String.Format("select EMail from vw_aspnet_MembershipUsers where UserName='{0}'", ds1.Tables[0].Rows[t]["UserName"].ToString()), ref obj, null);
                    if (obj != null)
                    {
                        try
                        {
                            MailAddress mTo = new MailAddress((string)obj);
                            mm.Bcc.Add(mTo);
                        }
                        catch { }
                    }
                }
                ds1.Clear();
                res = ExecuteQuery(String.Format("select id_card from V_Cards_StorageDocs where id_doc={0} and inform_production=1", id_doc), ref ds1, null);
                ArrayList al = new ArrayList();
                if (ds1.Tables.Count == 0)
                    return;
                for (int t = 0; t < ds1.Tables[0].Rows.Count; t++)
                {
                    if (ds1.Tables[0].Rows[t]["id_card"] != DBNull.Value)
                        al.Add(Convert.ToInt32(ds1.Tables[0].Rows[t]["id_card"]));
                }
                ds1.Clear();
                if (al.Count == 0)
                    return;
                SmtpClient sc = new SmtpClient(ConfigurationSettings.AppSettings["SmtpServer"]);
                sc.Credentials = new NetworkCredential(ConfigurationSettings.AppSettings["EMailFrom"], ConfigurationSettings.AppSettings["Pwd"]);
                MailAddress mailFrom = new MailAddress(ConfigurationSettings.AppSettings["EMailFrom"], "CardPerso");
                mm.From = mailFrom;
                mm.Subject = "CardPerso - информация о выпуске карт";
                mm.Body = "Выпущены следующие карты:\n";
                mm.Body += "<table width='100%' border='2'><tr><td>ФИО</td><td>PAN</td><td>Филиал заведения</td><td>Филиал выдачи</td></tr>";
                mm.IsBodyHtml = true;
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select pan, dateProd, fio, DepBranchInit, DepBranchCard from V_Cards where id=@id";
                comm.Parameters.Add("@id", SqlDbType.Int);
                for (int t = 0; t < al.Count; t++)
                {
                    ds1.Clear();
                    comm.Parameters["@id"].Value = (int)al[t];
                    res = ExecuteCommand(comm, ref ds1, null);
                    if (ds1.Tables.Count == 0)
                        continue;
                    string str = "<tr>";
                    str += "<td>" + ds1.Tables[0].Rows[0]["fio"].ToString().Trim() + "</td>";
                    str += "<td>" + ds1.Tables[0].Rows[0]["pan"].ToString().Trim() + "</td>";
                    //str += Convert.ToDateTime(ds1.Tables[0].Rows[0]["DateProd"]).ToShortDateString().PadRight(20) + " ";
                    str += "<td>" + ds1.Tables[0].Rows[0]["DepBranchInit"].ToString().Trim() + "</td>";
                    str += "<td>" + ds1.Tables[0].Rows[0]["DepBranchCard"].ToString().Trim() + "</td>";
                    str += "</tr>";
                    mm.Body += str;
                }
                mm.Body += "</table>";
                try
                {
                    if (mm.Bcc.Count > 0)
                    {
                        sc.Send(mm);
                        lbInform.Text = "Служебная записка отправлена по рассылке";
                    }
                }
                catch
                {
                }
            }
        }
        protected void bPanSearch_Click(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (tbPanSearch.Text.Trim().Length == 0 && tbFioSearch.Text.Trim().Length == 0)
                {
                    ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                    return;
                }
                object obj = null;
                DataSet ds = new DataSet();
                ExecuteScalar(String.Format("select count(*) from Cards_StorageDocs where id_doc={0}", Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"])), ref obj, null);
                int cnt = Convert.ToInt32(obj);
                string field = "V_Cards_StorageDocs.pan", val = tbPanSearch.Text.Trim();
                if (val.Length >= 16)
                {
                    field = "panhash";
                    SHA1Managed sha = new SHA1Managed();
                    val = Utilities.Bin2AHex(sha.ComputeHash(System.Text.Encoding.ASCII.GetBytes(tbPanSearch.Text.Trim())));
                }
                res = ExecuteQuery(String.Format("select *, cards.passport from V_Cards_StorageDocs join cards on cards.id=V_Cards_StorageDocs.id_card where id_doc={0} and {3} like '%{1}%' and V_Cards_StorageDocs.fio like '%{2}%'", Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), val, tbFioSearch.Text.Trim(), field), ref ds, null);

                lbCountCard.Text = String.Format("Карты: {0}/{1}", ds.Tables[0].Rows.Count, cnt);

                gvCards.DataSource = ds.Tables[0];
                gvCards.DataBind();



                /*            int i = 0;
                            for (i = gvCards.SelectedIndex; i < gvCards.Rows.Count; i++)
                            {
                                if (tbPanSearch.Text.Trim().Length > 0 && Convert.ToString(gvCards.DataKeys[i].Values["pan"]).IndexOf(tbPanSearch.Text.Trim()) >= 0)
                                {
                                    gvCards.SelectedIndex = i;
                                    gvCards.Rows[i].Focus();
                                    break;
                                }
                                if (tbFioSearch.Text.Trim().Length > 0 && Convert.ToString(gvCards.DataKeys[i].Values["fio"]).ToLower().IndexOf(tbFioSearch.Text.Trim().ToLower()) >= 0)
                                {
                                    gvCards.SelectedIndex = i;
                                    gvCards.Rows[i].Focus();
                                    break;
                                }
                            }
                            if (i == gvCards.Rows.Count)
                            {
                                for (i = 0; i < gvCards.SelectedIndex; i++)
                                {
                                    if (tbPanSearch.Text.Trim().Length > 0 && Convert.ToString(gvCards.DataKeys[i].Values["pan"]).IndexOf(tbPanSearch.Text.Trim()) >= 0)
                                    {
                                        gvCards.SelectedIndex = i;
                                        gvCards.Rows[i].Focus();
                                        break;
                                    }
                                    if (tbFioSearch.Text.Trim().Length > 0 && Convert.ToString(gvCards.DataKeys[i].Values["fio"]).ToLower().IndexOf(tbFioSearch.Text.Trim().ToLower()) >= 0)
                                    {
                                        gvCards.SelectedIndex = i;
                                        gvCards.Rows[i].Focus();
                                        break;
                                    }
                                }
                            }*/
                SetButtonCard();
            }
        }
        protected void bCardProperty_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                AddAutoStorageProduct(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]), null);
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                SetButtonCard();
            }
        }
        public string AddAutoStorageDoc(int id_doc, int id_type, int id_branch, int id_act, SqlTransaction trans)
        {
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    DataSet ds = new DataSet();
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;

                    da.SelectCommand.CommandText = "delete from Products_StorageDocs where id_doc=@id";
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.ExecuteNonQuery();

                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.CommandText = "delete from Cards_StorageDocs where id_doc=@id";
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.ExecuteNonQuery();

                    if (id_type == 1)
                        InsertProductsFromCards(id_doc, 1, -1, "new", trans);

                    if (id_type == 2)
                        InsertPins(id_doc, 1, -1, "new", trans);

                    if (id_type == 3)
                    {
                        InsertProductsFromStorage(id_doc, 1, "wrk", "perso", trans);
                        InsertCards(id_doc, 1, id_type, -1, trans);
                    }

                    if (id_type == 4)
                        InsertProductsFromStorage(id_doc, 2, "wrk", "perso", trans);

                    if (id_type == 5)
                    {
                        ExecuteQuery(String.Format("select distinct id_BranchCard from V_Cards where (id_stat=2) and ((id_BranchCard_parent={0}) or (id_BranchCard={0}))", id_branch.ToString()), ref ds, trans);
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            int id_branch_t = Convert.ToInt32(ds.Tables[0].Rows[i]["id_BranchCard"]);

                            //InsertProductsFromCards(id_doc, 2, id_branch_t, "perso");
                            //InsertPins(id_doc, 2, id_branch_t, "perso");
                            InsertCards(id_doc, 2, id_type, id_branch_t, trans);
                        }
                        InsertProductsFromCardsDoc(id_doc, 2, "perso", trans);
                        InsertDopProductsFromCardsDoc(id_doc, 2, "new", trans);
                        //InsertPinsDoc(id_doc, 2, "perso", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }

                    if (id_type == 6)
                    {
                        InsertFromDocs(id_doc, id_act, trans);
                        //ExecuteQuery(String.Format("select distinct id_BranchCard from V_Cards where (id_stat=3) and ((id_BranchCard_parent={0}) or (id_BranchCard={0}))", id_branch.ToString()), ref ds, null);
                        //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        //{
                        //    int id_branch_t = Convert.ToInt32(ds.Tables[0].Rows[i]["id_BranchCard"]);
                        //    InsertCards(id_doc, 3, id_type, id_branch_t);
                        //}
                    }

                    if (id_type == 7 && id_branch == -1)
                        InsertCards(id_doc, 4, id_type, -1, trans);

                    if (id_type == 8)
                    {
                        InsertCards(id_doc, 1, id_type, -1, trans);
                        InsertProductsFromCardsDoc(id_doc, 1, "perso", trans);
                        //InsertPinsDoc(id_doc, 1, "perso", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }

                    //if (id_type == 9)
                    //{
                    //    InsertCards(id_doc, 4, id_type, -1);
                    //}

                    if (id_type == 10)
                    {
                        //InsertCards(id_doc, 5, id_type, id_branch);
                        InsertFromDocs(id_doc, id_act, trans);
                    }

                    if (id_type == 11)
                    {
                        InsertProductsFromStorage(id_doc, -1, "brak", "brak", trans);
                        InsertCards(id_doc, 6, id_type, -1, trans);
                    }
                }
            
            return "";
        }
        public int CntCardsInStorageDocs(int id_doc, SqlTransaction trans)
        {
            int cnt = -1;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.CommandText = "select Count(id) as cnt from Cards_StorageDocs where id_doc=@id_doc";
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    cnt = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                }
            return cnt;
        }
        public int CntProductsInStorageDocs(int id_doc, string sost, int onlycards, SqlTransaction trans)
        {
            int cnt = -1;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.CommandText = "select sum(cnt_" + sost + ") as cnt from V_Products_StorageDocs where id_doc=@id_doc and (id_type=1 or id_type=2)";
                    if (onlycards == 1) //только карты
                        da.SelectCommand.CommandText = "select sum(cnt_" + sost + ") as cnt from V_Products_StorageDocs where id_doc=@id_doc and id_type=1";
                    if (onlycards == 2) //только расходники
                        da.SelectCommand.CommandText = "select sum(cnt_" + sost + ") as cnt from V_Products_StorageDocs where id_doc=@id_doc and id_type=4";
                        
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    object obj = da.SelectCommand.ExecuteScalar();
                    cnt = (obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                }
            return cnt;
        }
        public int CntStorage(int id_prb, string sost, SqlTransaction trans)
        {
            int cnt = -1;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    if (sost == "new")
                    {
    //                    da.SelectCommand.CommandText = "select parent from V_Storage where id_prb=@id_prb";
                        da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prb;
  //                      object pid = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                        da.SelectCommand.CommandText = "select sum(cnt_" + sost + ") from V_Storage where id_prb=@id_prb or id_prod in (select parent from V_ProductsBanks where id_prb=@id_prb)";
//                        da.SelectCommand.Parameters.Add("@id_par", SqlDbType.Int).Value = (pid == DBNull.Value) ? -100 : (int)pid;
                        object pid = da.SelectCommand.ExecuteScalar();
                        cnt = (pid == DBNull.Value ) ? 0 : Convert.ToInt32(pid);
                    }
                    else
                    {
                        da.SelectCommand.CommandText = "select cnt_" + sost + " from Storage where id_prb=@id_prb";
                        da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prb;
                        cnt = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                    }
                }
            return cnt;
        }
        public bool CheckStatusCardsInStorageDocs(int id_stat, int id_doc, SqlTransaction trans)
        {
            bool res = true;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.CommandText = "select Count(Cards_StorageDocs.id) as cnt from Cards_StorageDocs inner join Cards on Cards_StorageDocs.id_card = Cards.id where (Cards.id_stat <> @id_stat) and (Cards_StorageDocs.id_doc=@id_doc)";
                    da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    int cnt = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                    if (cnt > 0) return false;
                }

            return res;
        }
        public bool CheckPropCardsInStorageDocs(int id_prop, int id_doc, SqlTransaction trans)
        {
            bool res = true;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                da.SelectCommand.CommandText = "select Count(Cards_StorageDocs.id) as cnt from Cards_StorageDocs inner join Cards on Cards_StorageDocs.id_card = Cards.id where (Cards.id_prop <> @id_prop) and (Cards_StorageDocs.id_doc=@id_doc)";
                da.SelectCommand.Parameters.Add("@id_prop", SqlDbType.Int).Value = id_prop;
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                int cnt = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                if (cnt > 0) return false;
            }

            return res;
        }

        public bool CheckStatusCardsInStorageDocsPerso(int id_doc, SqlTransaction trans)
        {
            bool res = true;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                da.SelectCommand.CommandText = "select Count(Cards_StorageDocs.id) as cnt from Cards_StorageDocs inner join Cards on Cards_StorageDocs.id_card = Cards.id where (Cards.id_stat <> @id_stat1 and Cards.id_stat <> @id_stat2) and (Cards_StorageDocs.id_doc=@id_doc)";
                da.SelectCommand.Parameters.Add("@id_stat1", SqlDbType.Int).Value = 2;
                da.SelectCommand.Parameters.Add("@id_stat2", SqlDbType.Int).Value = 9;
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                int cnt = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                if (cnt > 0) return false;
            }

            return res;
        }
        public bool CheckCntCardsProducts(int id_doc, string sost, SqlTransaction trans)
        {
            bool res = true;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    DataSet ds = new DataSet();
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_doc={0}) group by id_prb", id_doc.ToString()), ref ds, trans);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        int cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                        int id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);

                        if (cnt != CntProductInStorageDocs(id_doc, id_prb, sost, trans))
                            return false;
                    }
                    if (CntCardsInStorageDocs(id_doc, trans) != CntProductsInStorageDocs(id_doc, sost, 1, trans))
                        return false;
                }
            return res;
        }
        public bool CheckCntCardsProductsReturnToBank(int id_doc, SqlTransaction trans)
        {
            bool res = true;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    DataSet ds = new DataSet();
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_doc={0}) and id_prop<>10 group by id_prb", id_doc.ToString()), ref ds, trans);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        int cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                        int id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                        if (cnt != CntProductInStorageDocs(id_doc, id_prb, "brak", trans))
                            return false;
                    }
                    ds.Clear();
                    ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_doc={0}) and id_prop=10 group by id_prb", id_doc.ToString()), ref ds, trans);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        int cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                        int id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                        if (cnt != CntProductInStorageDocs(id_doc, id_prb, "perso", trans))
                            return false;
                    }
                    //                        if (CntCardsInStorageDocs(id_doc) != CntProductsInStorageDocs(id_doc, sost))
                    //                            return false;
                }
            return res;
        }
        public bool CheckStatusCardsInStorageDocsReturnToBank(int id_doc, SqlTransaction trans)
        {
            bool res = true;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    int st = 2;
                    if (branch_main_filial > 0) st = 4;
                    da.SelectCommand.CommandText = String.Format("select Count(Cards_StorageDocs.id) as cnt from Cards_StorageDocs inner join Cards on Cards_StorageDocs.id_card = Cards.id where (Cards.id_stat <> 6 and Cards.id_stat <> {0}) and (Cards_StorageDocs.id_doc=@id_doc)",st);
                    //                        da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    int cnt = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                    if (cnt > 0) return false;
                }
            return res;
        }
        public string StorageM(int cnt, int id_prod, string sost, SqlTransaction trans)
        {
            if (isMainFilial()) return "";

            if (cnt == 0)
                return "";
            string res = String.Empty;
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    if (sost == "new")
                    {
                        da.SelectCommand.CommandText = "select parent from V_ProductsBanks where id_prb=@id_prod";
                        da.SelectCommand.Parameters.Add("@id_prod", SqlDbType.Int).Value = id_prod;
                        object pid = da.SelectCommand.ExecuteScalar();
                        if (pid == DBNull.Value || pid == null)
                        {
                            da.SelectCommand.CommandText = "update Storage set cnt_" + sost + "=cnt_" + sost + "-@cnt where id_prb=@id_prod";
                            da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                            da.SelectCommand.ExecuteNonQuery();
                        }
                        else
                        {
                            da.SelectCommand.CommandText = "select id_prb from V_Storage where id_prod=@id_prod";
                            da.SelectCommand.Parameters["@id_prod"].Value = pid;
                            pid = da.SelectCommand.ExecuteScalar();
                            da.SelectCommand.CommandText = "update Storage set cnt_" + sost + "=cnt_" + sost + "-@cnt where id_prb=@id_prod";
                            da.SelectCommand.Parameters["@id_prod"].Value = Convert.ToInt32(pid);
                            da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                            da.SelectCommand.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        da.SelectCommand.CommandText = "update Storage set cnt_" + sost + "=cnt_" + sost + "-@cnt where id_prb=@id_prod";
                        da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                        da.SelectCommand.Parameters.Add("@id_prod", SqlDbType.Int).Value = id_prod;
                        da.SelectCommand.ExecuteNonQuery();
                    }                    
                }
            return res;
        }
        
        public string StorageP(int cnt, int id_prod, string sost, SqlTransaction trans)
        {

            if (isMainFilial()) return "";

            if (cnt == 0)
                return "";
            //if (sost != "new") return "";
            string res = String.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.Parameters.Add("@id_prod", SqlDbType.Int).Value = id_prod;
                    if (sost == "new")
                    {
                        da.SelectCommand.CommandText = "select parent from V_ProductsBanks where id_prb=@id_prod";
                        object pid = da.SelectCommand.ExecuteScalar();
                        if (pid != DBNull.Value)
                        {
                            da.SelectCommand.CommandText = "select id_prb from V_Storage where id_prod=@id_prod";
                            da.SelectCommand.Parameters["@id_prod"].Value = Convert.ToInt32(pid);
                            id_prod = Convert.ToInt32(da.SelectCommand.ExecuteScalar());
                        }
                    }
                    da.SelectCommand.Parameters["@id_prod"].Value = id_prod;
                    da.SelectCommand.CommandText = "select id from Storage where id_prb=@id_prod";
                    
                    object obj = da.SelectCommand.ExecuteScalar();
                    da.SelectCommand.Parameters.Clear();
                    if (Convert.ToInt32(obj) > 0)
                    {
                        da.SelectCommand.CommandText = "update Storage set cnt_" + sost + "=cnt_" + sost + "+@cnt where id_prb=@id_prod";
                        da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                        da.SelectCommand.Parameters.Add("@id_prod", SqlDbType.Int).Value = id_prod;
                        da.SelectCommand.ExecuteNonQuery();
                    }
                    else
                    {
                        da.SelectCommand.CommandText = "insert into Storage (id_prb,cnt_" + sost + ") values (@id_prb,@cnt)";
                        da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prod;
                        da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                        da.SelectCommand.ExecuteNonQuery();
                    }
                }
            return res;
        }
        public string UpdateCardsDate(int id_doc, int id_stat, int sost, string date_name, DateTime dt, string userGuid, SqlTransaction trans)
        {
            //сделан чтобы после возрата с упаковки правильно выставить статус карты
            return UpdateCardsDate(id_doc, id_stat, sost, date_name, dt, userGuid, false, false, null);
        }
        public string UpdateCardsDate(int id_doc, int id_stat, int sost, string date_name, DateTime dt, string userGuid, bool afterWrapping, SqlTransaction trans)
        {
            //сделан чтобы после возрата с упаковки правильно выставить статус карты
            return UpdateCardsDate(id_doc, id_stat, sost, date_name, dt, userGuid, afterWrapping, false, null);
        }
        public string UpdateCardsDate(int id_doc, int id_stat, int sost, string date_name, DateTime dt, string userGuid, bool afterWrapping, bool dontReceive, SqlTransaction trans)
        {
            string res = String.Empty;
            WebLog.LogClass.WriteToLog("StorDoc.UpdateCardsDate Start id_doc={0}, id_stat = {1}, sost = {2}, date_name = {3}, user = {4}, userbranch = {5}", id_doc, id_stat, sost, date_name, User.Identity.Name, sc.BranchId(User.Identity.Name));                                               
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.CommandText = "select count(id) from Cards_StorageDocs where id_doc = @id_doc";
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    object co = da.SelectCommand.ExecuteScalar();
                    int count = 0;
                    if (co != DBNull.Value)
                    {
                        count = Convert.ToInt32(co);
                    }
                    WebLog.LogClass.WriteToLog("StorDoc.UpdateCardsDate id_doc={0}, count card={1}", id_doc, count);
                    
                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat," + date_name + "=@dt where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    if (id_stat == 8 && sost == 0)
                    {
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat," + date_name + "=@dt, idWorker=@idWorker where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                        da.SelectCommand.Parameters.Add("@idWorker", SqlDbType.UniqueIdentifier).Value = new Guid(userGuid);
                    }
                    if (id_stat == 4 && sost == 1)
                    {
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat," + date_name + "=@dt, idWorker=@idWorker where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                        da.SelectCommand.Parameters.Add("@idWorker", SqlDbType.UniqueIdentifier).Value = DBNull.Value;
                    }
                    if (id_stat == 4 && sost == 0)
                    {
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat," + date_name + "=@dt, id_BranchCurrent=@branchCurrent where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                        int id_branch = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id_branch"]);
                        da.SelectCommand.Parameters.Add("@branchCurrent", SqlDbType.Int).Value = id_branch;
                    }
                    if (id_stat == 3 && sost == 1)
                    {
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat," + date_name + "=@dt, id_BranchCurrent=@branchCurrent where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                        da.SelectCommand.Parameters.Add("@branchCurrent", SqlDbType.Int).Value = DBNull.Value;
                    }
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                    if (sost == 0)
                    {
                        if (dt == DateTime.MinValue)
                            da.SelectCommand.Parameters.Add("@dt", SqlDbType.DateTime).Value = DateTime.Today;
                        else
                            da.SelectCommand.Parameters.Add("@dt", SqlDbType.DateTime).Value = dt;
                    }
                    else
                        da.SelectCommand.Parameters.Add("@dt", SqlDbType.DateTime).Value = DBNull.Value;
                    da.SelectCommand.ExecuteNonQuery();
                    // Отдельно меняем статус на хранилище (упаковка) для карт после персонализации, которые требуют упаковки
                    if (id_stat == 2 && sost == 0 && !afterWrapping)
                    {
                        da.SelectCommand.CommandText = "update Cards set id_stat=@id_stat where (id in (select id_card from V_Cards_StorageDocs where id_doc=@id_doc and wrapping=1))";
                        da.SelectCommand.Parameters["@id_stat"].Value = 9;
                        da.SelectCommand.ExecuteNonQuery();
                    }
                    // при отмене отправки на упаковку, меняем на соответствующий статус тех карт, которые требуется
                    if (id_stat == 2 && sost == 1 && afterWrapping)
                    {
                        da.SelectCommand.CommandText = "update Cards set id_stat=@id_stat where (id in (select id_card from V_Cards_StorageDocs where id_doc=@id_doc and wrapping=1))";
                        da.SelectCommand.Parameters["@id_stat"].Value = 9;
                        da.SelectCommand.ExecuteNonQuery();
                    }
                    // при отказе приема в филиал и отправке обратно в банк на уничтожение меняем свойство карты на 11
                    if (id_stat == 5 && sost == 0 && dontReceive)
                    {
                        da.SelectCommand.CommandText = "update Cards set id_prop=@id_stat where (id in (select id_card from V_Cards_StorageDocs where id_doc=@id_doc))";
                        da.SelectCommand.Parameters["@id_stat"].Value = 11;
                        da.SelectCommand.ExecuteNonQuery();
                    }
                    // при отмене документа отказа приема в филиал и отправке обратно в банк на уничтожение, меняем свойство карты на хорошее
                    if (id_stat == 3 && sost == 1 && dontReceive)
                    {
                        da.SelectCommand.CommandText = "update Cards set id_prop=@id_stat where (id in (select id_card from V_Cards_StorageDocs where id_doc=@id_doc))";
                        da.SelectCommand.Parameters["@id_stat"].Value = 1;
                        da.SelectCommand.ExecuteNonQuery();
                    }
                    // при отмене документа выдачи клиенту, смотрим не была ли эта выдачи из книги 124 и тогда ставим соответствующий статутс.
                    if (id_stat == 4 && sost == 1)
                    {
                        da.SelectCommand.CommandText = "update Cards set id_stat=19 where (id in (select id_card from V_Cards_StorageDocs where id_doc=@id_doc and id_person is not NULL))";
                        da.SelectCommand.ExecuteNonQuery();
                    }
                }
                WebLog.LogClass.WriteToLog("StorDoc.UpdateCardsDate End id_doc={0}, id_stat = {1}, sost = {2}, date_name = {3}, user = {4}, userbranch = {5}", id_doc, id_stat, sost, date_name, User.Identity.Name, sc.BranchId(User.Identity.Name));                                               
            return res;
        }
        public string UpdateCardsCourier(int id_doc, int id_stat, int sost, DateTime dt, string invoice, string id_courier, string courier_name, SqlTransaction trans)
        {
            string res = String.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat,id_courier=@id_courier,invoice=@invoice,courier_name=@courier_name,date_courier=@date_courier where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                    da.SelectCommand.Parameters.Add("@invoice", SqlDbType.VarChar, 20).Value = invoice;
                    da.SelectCommand.Parameters.Add("@courier_name", SqlDbType.VarChar, 50).Value = courier_name;
                    if (id_courier == "")
                        da.SelectCommand.Parameters.Add("@id_courier", SqlDbType.Int).Value = DBNull.Value;
                    else
                        da.SelectCommand.Parameters.Add("@id_courier", SqlDbType.Int).Value = Convert.ToInt32(id_courier);
                    if (sost == 0)
                    {
                        if (dt == DateTime.MinValue)
                            da.SelectCommand.Parameters.Add("@date_courier", SqlDbType.DateTime).Value = DateTime.Today;
                        else
                            da.SelectCommand.Parameters.Add("@date_courier", SqlDbType.DateTime).Value = dt;
                    }
                    else
                        da.SelectCommand.Parameters.Add("@date_courier", SqlDbType.DateTime).Value = DBNull.Value;
                    da.SelectCommand.ExecuteNonQuery();
                }
            return res;
        }        
        /// <summary>
        /// изменение статуса при возврате карт клиентом в филиал 
        /// </summary>
        /// <param name="id_doc">id документа</param>
        /// <param name="id_stat">новый статус</param>
        /// <param name="sost">изменение свойства для возврата карт в филиал (для остальных ставим 0)</param>
        /// <param name="trans"></param>
        /// <returns></returns>
        public string UpdateCardsDoc(int id_doc, int id_stat, int sost, SqlTransaction trans)
        {
            string res = String.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat, id_prop=@id_prop where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                    da.SelectCommand.Parameters.Add("@id_prop", SqlDbType.Int).Value = (sost == 1) ? 8 : 1;
                    da.SelectCommand.ExecuteNonQuery();
                }
            return res;
        }
        public string UpdateCardsDocPodotchet(int id_doc, int id_stat, int sost, SqlTransaction trans)
        {
            string res = String.Empty;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                da.SelectCommand.Parameters.Clear();
                if (sost == 0)
                    da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat, id_branchCurrent=" + current_branch_id.ToString() + ", id_person=@id_person where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                else
                    da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat, id_branchCurrent=id_branchcard, id_person=@id_person where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                da.SelectCommand.Parameters.Add("@id_person", SqlDbType.Int).Value = DBNull.Value;
                // если карта оказывается у подотчетного лица, то в id_person проставляем его id, чтобы понять, у кого именно она сейчас числится
                // 15 - это подотчет
                // 19 - это выдача по 124 книге, по сути тоже самое
                if (id_stat == 19)
                    da.SelectCommand.Parameters["@id_person"].Value = Convert.ToInt32(Session["CurrentUserId"]);

                //11.12.2019 раньше был отдельный прием подотчетным лицом карт. Теперь после выдачи они сразу числятся на нем и поэтому они принимались в рамках прошлого if, поскольку current user был подотчетным
                //а теперь надо проставлять того, кому предназначены карты
                if (id_stat == 15)
                {
                    SqlCommand comm = conn.CreateCommand();
                    comm.CommandText = $"select id_person from [AccountablePerson_StorageDocs] where id_doc={id_doc}";
                    try
                    {
                        da.SelectCommand.Parameters["@id_person"].Value = Convert.ToInt32(comm.ExecuteScalar());
                    }
                    catch
                    {
                        return res;
                    }
                }
                da.SelectCommand.ExecuteNonQuery();
            }
            return res;
        }


        public string UpdateCardsDocExpertiza(int id_doc, int id_stat, int sost, SqlTransaction trans)
        {
            string res = String.Empty;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                da.SelectCommand.Parameters.Clear();
                if (sost == 0)
                {
                    if (id_stat != 130)
                    {
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat, id_branchCurrent=" + current_branch_id.ToString() + " where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    }
                    else // подтверждение экспертизы
                    {
                        da.SelectCommand.CommandText = "Update Cards set id_stat=(case when id_prop = 1 then 2 else 6 end), id_branchCurrent=" + current_branch_id.ToString() + ",id_branchCard=" + id_branchExpertiza.ToString() + " where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    }
                }
                else
                    da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat, id_prop=1, id_branchCurrent=id_branchcard where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) update Cards_StorageDocs set resultexpertiza=NULL where id_doc=@id_doc";
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                da.SelectCommand.ExecuteNonQuery();
            }
            return res;
        }
        public string UpdateCardsTerminate(int id_doc, int id_stat, int sost, DateTime dt, SqlTransaction trans)
        {
            string res = String.Empty;
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    if (sost == 0)
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat,dateSendTerminate=@dt,comment=(select top 1 comment from Cards_StorageDocs where id_card=Cards.id and id_doc=@id_doc) where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop<>10";
                    else
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat,dateSendTerminate=@dt,comment='' where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc))";
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.Parameters.Add("@id_stat", SqlDbType.Int).Value = id_stat;
                    if (sost == 0)
                    {
                        if (dt == DateTime.MinValue)
                            da.SelectCommand.Parameters.Add("@dt", SqlDbType.DateTime).Value = DateTime.Today;
                        else
                            da.SelectCommand.Parameters.Add("@dt", SqlDbType.DateTime).Value = dt;
                    }
                    else
                        da.SelectCommand.Parameters.Add("@dt", SqlDbType.DateTime).Value = DBNull.Value;
                    da.SelectCommand.ExecuteNonQuery();
                    if (sost == 0)
                    {
                        da.SelectCommand.CommandText = "Update Cards set id_stat=@id_stat,comment=(select top 1 comment from Cards_StorageDocs where id_card=Cards.id and id_doc=@id_doc) where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop=10";
                        da.SelectCommand.ExecuteNonQuery();
                    }
                }
            return res;
        }
        public string UpdateCardsDateReturnToBank(int id_doc, int sost, DateTime dt, string userGuid, SqlTransaction trans)
        {
            string res = String.Empty;
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                SqlCommand comm = new SqlCommand("", conn);
                if (trans != null)
                    comm.Transaction = trans;
                comm.CommandTimeout = conn.ConnectionTimeout;
                comm.Parameters.Clear();
                comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                comm.Parameters.Add("@dt", SqlDbType.DateTime);
                if (sost == 1) // отмена подтверждения
                {
                    comm.Parameters["@dt"].Value = DBNull.Value;
                    //те которые не 'неверный филиал отправки' - в данный момент эти карты оказываются со статусом CardOK
                    //comm.CommandText = "Update Cards set id_stat=5, dateGetTerminate=@dt where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop <> 1";
                    comm.CommandText = "Update Cards set id_stat=5, dateGetTerminate=@dt,id_BranchCurrent=id_BranchCard where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop <> 1";
                    comm.ExecuteNonQuery();
                    // те которые 'неверный филиал отправки'
                    //comm.CommandText = "Update Cards set id_stat=5, id_prop=10 where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop = 1";
                    comm.CommandText = "Update Cards set id_stat=5, id_prop=10,id_BranchCurrent=id_BranchCard where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop = 1";
                    comm.ExecuteNonQuery();
                }
                else //подтверждение
                {
                    if (dt == DateTime.MinValue)
                        comm.Parameters["@dt"].Value = DateTime.Today;
                    else
                        comm.Parameters["@dt"].Value = dt;
                    //те которые не 'неверный филиал отправки'

                    //13.04.20 если это работник персобюро то таже если он залогинин в филиале, то принимает в хранилище пц (есть псевдо-филиал пц)
                    int newbranch = (sc.UserAction(User.Identity.Name, Restrictions.Perso)) ? 0 : current_branch_id;

                    //!!!---------------------------------------
                    //comm.CommandText = "Update Cards set id_stat=6, dateGetTerminate=@dt where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop <> 10";
                    //~!if(branch_main_filial>0) comm.CommandText = "Update Cards set id_stat=6, id_BranchCurrent=" + branch_main_filial.ToString() + ", dateGetTerminate=@dt where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop <> 10";
                    //~!else 
                    comm.CommandText = "Update Cards set id_stat=6, id_BranchCurrent=" + newbranch.ToString() + ", dateGetTerminate=@dt where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop <> 10"; 
                    //------------------------------------------
                    comm.ExecuteNonQuery();
                    //выбираем 'неверный филиал отправки'
                    comm.Parameters.Clear();
                    // нижнюю строчку закоментарил, чтобы свойство карты осталось 'неверный филиал отправки' и она не уходила с близжайшей рассылкой
//                    comm.CommandText = "Update Cards set id_stat=2, date_courier=null, dateReceipt=null, id_prop=1 where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop = 10";
                    //comm.CommandText = "Update Cards set id_stat=2, date_courier=null, dateReceipt=null, dateGetTerminate=null where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop = 10";
                    //~!if (branch_main_filial > 0) comm.CommandText = "Update Cards set id_stat=4, id_BranchCurrent=" + branch_main_filial.ToString() + ", date_courier=null, dateReceipt=null, dateGetTerminate=null where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop = 10";
                    //~!else 
                    comm.CommandText = "Update Cards set id_stat=2, id_BranchCurrent=" + newbranch.ToString() + ", date_courier=null, dateReceipt=null, dateGetTerminate=null where (id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)) and id_prop = 10";
                    comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    comm.ExecuteNonQuery();
                }
            return res;
        }
        public string UpdateCardsFilialFilial(int id_doc, int sost, SqlTransaction trans)
        {
            string ret = "";
                SqlCommand comm = conn.CreateCommand();
                DataSet ds = new DataSet();
                ret = ExecuteQuery("select id_branch, id_courier, id_act from StorageDocs where id=" + id_doc.ToString(), ref ds, trans);
                if (ds.Tables[0].Rows.Count == 0)
                {
                    return ret;
                }
                int newbr = Convert.ToInt32(ds.Tables[0].Rows[0]["id_courier"]);
                int oldbr = Convert.ToInt32(ds.Tables[0].Rows[0]["id_branch"]);
                bool isForKilling = (Convert.ToInt32(ds.Tables[0].Rows[0]["id_branch"]) == 1);
                comm.CommandText = String.Format("update cards set id_stat=@stat, id_branchcard=@branch, datereceipt=@dateR where id in (select id_card from Cards_StorageDocs where id_doc=@id_doc)");
                comm.Transaction = trans;
                comm.Parameters.Add("@stat", SqlDbType.Int).Value = (sost == 0) ? 3 : 4;
                comm.Parameters.Add("@branch", SqlDbType.Int);
                comm.Parameters.Add("@dateR", SqlDbType.DateTime);
                comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                if (sost == 0) // подтверждение
                {
                    comm.Parameters["@branch"].Value = newbr;
                    comm.Parameters["@dateR"].Value = DBNull.Value;
                }
                else
                {
                    comm.Parameters["@branch"].Value = oldbr;
                    comm.Parameters["@dateR"].Value = DateTime.Now.Date;
                }
                comm.ExecuteNonQuery();
            return ret;
        }
        public string GetFilialPeople(int id_fil, SqlTransaction trans)
        {
            if (id_fil < 0)
                return "";
            string fio = "";
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.CommandText = "select adress+'; '+people as fio from Branchs where id=@id";
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id_fil;
                    fio = da.SelectCommand.ExecuteScalar().ToString();
                }
            return fio;
        }
        public string ExecuteNonQuery(SqlCommand comm, SqlTransaction transaction)
        {
            string res = String.Empty;
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                comm.Connection = conn;
                comm.CommandTimeout = conn.ConnectionTimeout;
                if (transaction != null)
                    comm.Transaction = transaction;
                comm.ExecuteNonQuery();
            return res;
        }
        public string ExecuteNonQuery(String query, SqlTransaction transaction)
        {
            string res = String.Empty;
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                SqlCommand comm = conn.CreateCommand();
                comm.CommandTimeout = conn.ConnectionTimeout;
                comm.CommandText = query;
                if (transaction != null)
                    comm.Transaction = transaction;
                comm.ExecuteNonQuery();
            return res;
        }
        public string  ExecuteCommand(SqlCommand comm, ref DataSet ds, SqlTransaction transaction)
        {
            string res = "";
                comm.Connection = conn;
                SqlDataAdapter da = new SqlDataAdapter();
                if (transaction != null)
                    comm.Transaction = transaction;
                da.SelectCommand = comm;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                try
                {
                    da.Fill(ds);
                }
                catch
                {
                }
            return res;
        }
        public string AddAutoStorageProduct(int id_doc, int id_type, SqlTransaction trans)
        {
            string res = String.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    DataSet ds = new DataSet();
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;

                    da.SelectCommand.CommandText = "delete from Products_StorageDocs where id_doc=@id";
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.ExecuteNonQuery();

                    //if (id_type == 1)
                    //    InsertProductsFromCards(id_doc, 1, -1, "new");

                    //if (id_type == 2)
                    //    InsertPins(id_doc, 1, -1, "new");

                    //if (id_type == 3)
                    //{
                    //    InsertProductsFromStorage(id_doc, 1, "wrk", "perso");
                    //    InsertCards(id_doc, 1, id_type, -1);
                    //}

                    //if (id_type == 4)
                    //    InsertProductsFromStorage(id_doc, 2, "wrk", "perso");

                    if (id_type == 5)
                    {
                        InsertProductsFromCardsDoc(id_doc, 2, "perso", trans);
                        InsertDopProductsFromCardsDoc(id_doc, 2, "new", trans);
                        //InsertPinsDoc(id_doc, 2, "perso", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }

                    //if (id_type == 6)
                    //{
                    //    InsertProductsFromCardsDoc(id_doc, 4, "perso");
                    //}

                    if (id_type == 6)
                    {
                        InsertProductsFromCardsDoc(id_doc, 3, "perso", trans);
                        InsertDopProductsFromCardsDoc(id_doc, 3, "new", trans);
                        //InsertPinsDoc(id_doc, 3, "perso", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }

                    if (id_type == 8)
                    {
                        InsertProductsFromCardsDoc(id_doc, 1, "perso", trans);
                        //InsertPinsDoc(id_doc, 1, "perso", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }
                    // возврат ценностей из филиала в банк
                    if (id_type == 9)
                    {
                        InsertProductsFromCardsDocReturnToBank(id_doc, 4, trans);
                    }
                    if (id_type == 10)
                    {
                        InsertProductsFromCardsDocReturnToBank(id_doc, (int)CardStatus.CourierBank, trans);
                    }

                    if (id_type == 20)
                    {
                        InsertProductsFromCardsDoc(id_doc, 8, "perso", trans);
                    }

                    if (id_type == 23)
                    {
                        InsertProductsFromCardsExpertiza(id_doc, 13, trans);
                    }
                    // пересылка филиал-филиал
                    if (id_type == 13)
                    {
                        InsertProductsFromCardsDoc(id_doc, 4, "perso", trans);
                        InsertDopProductsFromCardsDoc(id_doc, 4, "new", trans);
                        InsertPinsDocPerso(id_doc, trans);
                        //InsertPinsDoc(id_doc, 4, "perso", trans);
                    }
                    //возврат с упаковки
                    if (id_type == 17)
                    {
                        InsertProductsFromCardsDoc(id_doc, 10, "perso", trans);
                        InsertDopProductsFromCardsDoc(id_doc, 10, "new", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }
                    //отправка на упаковку
                    if (id_type == 16)
                    {
                        InsertProductsFromCardsDoc(id_doc, 9, "perso", trans);
                        InsertDopProductsFromCardsDoc(id_doc, 9, "new", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }
                    // 4 документа по книге 124
                    if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124
                        || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124)
                    {
                        InsertProductsFromCardsDoc(id_doc, (int)FuncClass.CardStatusInDoc_NotConfirmed(id_type), "perso", trans);
                        InsertDopProductsFromCardsDoc(id_doc, (int)FuncClass.CardStatusInDoc_NotConfirmed(id_type), "new", trans);
                        InsertPinsDocPerso(id_doc, trans);
                    }
                // 4 документа по гоз
                if (id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz
                    || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz)
                {
                    InsertProductsFromCardsDoc(id_doc, (int)FuncClass.CardStatusInDoc_NotConfirmed(id_type), "perso", trans);
                    InsertDopProductsFromCardsDoc(id_doc, (int)FuncClass.CardStatusInDoc_NotConfirmed(id_type), "new", trans);
                    InsertPinsDocPerso(id_doc, trans);
                }
                // 4 документа по гоз-подотчет
                if (id_type == (int)TypeDoc.ToGozFromPodotchet || id_type == (int)TypeDoc.FromGozToPodotchet
                                                  || id_type == (int)TypeDoc.ToPodotchetFromGoz || id_type == (int)TypeDoc.FromPodotchetToGoz)
                {
                    InsertProductsFromCardsDoc(id_doc, (int)FuncClass.CardStatusInDoc_NotConfirmed(id_type), "perso", trans);
                    InsertDopProductsFromCardsDoc(id_doc, (int)FuncClass.CardStatusInDoc_NotConfirmed(id_type), "new", trans);
                    InsertPinsDocPerso(id_doc, trans);
                }


                //if (id_type == 10)
                //{
                //    InsertCards(id_doc, 5, id_type, -1);
                //}


            }
            ViewProducts(id_doc, 0);
            return res;
        }
        public string InsertProductsFromCards(int id_doc, int id_stat, int id_branch, string sost, SqlTransaction trans)
        {
            string res = String.Empty;
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                        string str_branch = "";
                        if (id_branch != -1)
                            str_branch = String.Format(" and (id_BranchCard={0})", id_branch.ToString());

                        ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from Cards where (id_stat={0}){1} group by id_prb", id_stat.ToString(), str_branch), ref ds, trans);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), sost, trans);
                        }
                    }
            return res;
        }
        public string InsertPins(int id_doc, int id_stat, int id_branch, string sost, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        ExecuteQuery(String.Format("dbo.PinCount {0},{1}", id_stat.ToString(), id_branch.ToString()), ref ds, trans);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), sost, trans);
                        }
                    }
            return res;
        }
        public string InsertProductsFromStorage(int id_doc, int id_type, string sost_in, string sost_out, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        string str_type = "";
                        if (id_type != -1)
                            str_type = String.Format("and (V_ProductsBanks_T.id_type = {0})", id_type.ToString());

                        ExecuteQuery(String.Format("select Storage.id_prb, Storage.cnt_" + sost_in + " as cnt from Storage inner join V_ProductsBanks_T on Storage.id_prb = V_ProductsBanks_T.id_prb where (Storage.cnt_" + sost_in + " > 0) {0}", str_type), ref ds, trans);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), sost_out, trans);
                        }
                    }
            return res;
        }
        public string InsertCards(int id_doc, int id_stat, int id_type, int id_branch, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;

                        string str_branch = "";
                        string str_prop = "";
                        if (id_branch != -1)
                            str_branch = String.Format(" and (id_BranchCard={0})", id_branch.ToString());
                        if (id_type == 8 || id_type == 5 || id_type == 6 || id_type == 7)
                            str_prop = " and id_prop = 1";
                        if (id_type == 9 || id_type == 10 || id_type == 11 || id_type == 12)
                            str_prop = " and id_prop <> 1";

                        //                        ExecuteQuery(String.Format("select id from Cards where id_stat = {0} and (id not in (select id_card from V_CardsTypeDocs where type={1})){2} order by pan", id_stat.ToString(), id_type.ToString(),str_branch), ref ds, null);
                        ExecuteQuery(String.Format("select id from Cards where id_stat = {0} and (id not in (select id_card from V_CardsTypeDocs where type={1})){2}{3} order by pan", id_stat.ToString(), id_type.ToString(), str_branch, str_prop), ref ds, trans);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                            da.SelectCommand.Parameters.Clear();
                            da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card) values (@id_doc,@id_card)";
                            da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                            da.SelectCommand.Parameters.Add("@id_card", SqlDbType.Int).Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);
                            da.SelectCommand.ExecuteNonQuery();
                        }
                    }
            return res;
        }
        public string InsertProductsFromCardsDoc(int id_doc, int id_stat, string sost, SqlTransaction trans)
        {
            string res = String.Empty;
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_stat={0}) and (id_doc={1}) group by id_prb", id_stat.ToString(), id_doc.ToString()), ref ds, trans);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), sost, trans);
                        }
                    }
            return res;
        }
        public string InsertDopProductsFromCardsDoc(int id_doc, int id_stat, string sost, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        ExecuteQuery(String.Format("select id_prb, cnt,(select Count(id_prb) from V_Cards_StorageDocs where id_prb=gg.id_prb_parent and id_stat={0} and id_doc={1}) as cnt_card from Products_Attachments gg where (id_prb_parent in (select id_prb from V_Cards_StorageDocs where id_stat={0} and id_doc={1}))", id_stat.ToString(), id_doc.ToString()), ref ds, trans);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            int cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]) * Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_card"]);
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), cnt, sost, trans);
                        }
                    }
            return res;
        }
        public string InsertPinsDoc1(int id_doc, int id_stat, string sost, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;

                        ExecuteQuery(String.Format("dbo.PinCountDoc {0},{1}", id_doc.ToString(), id_stat.ToString()), ref ds, trans);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), sost, trans);
                        }
                    }
            return res;
        }
        public string InsertPinsDocPerso(int id_doc, SqlTransaction trans)
        {
            string res = String.Empty;
            lock (Database.lockObjectDB)
            {
                SqlCommand comm = conn.CreateCommand();
                comm.Transaction = trans;
                comm.CommandText = "select id_bank, isPin from V_Cards_StorageDocs where id_doc=@id_doc";
                comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                SqlDataReader dr = comm.ExecuteReader();
                Hashtable ht = new Hashtable();
                while (dr.Read())
                {
                    if (dr["id_bank"] == DBNull.Value)
                        continue;
                    if (ht.ContainsKey(Convert.ToInt32(dr["id_bank"])))
                    {
                        if (Convert.ToBoolean(dr["isPin"]))
                            ht[Convert.ToInt32(dr["id_bank"])] = (int) ht[Convert.ToInt32(dr["id_bank"])] + 1;
                    }
                    else if (Convert.ToBoolean(dr["isPin"]))
                        ht.Add(Convert.ToInt32(dr["id_bank"]), (int) 1);
                }

                dr.Close();
                comm.Parameters.Clear();
                comm.CommandText = "select id_prb from V_ProductsBanks_T where id_type=2 and id_bank=@bank";
                comm.Parameters.Add("@bank", SqlDbType.Int);
                foreach (DictionaryEntry de in ht)
                {
                    comm.Parameters["@bank"].Value = (int) de.Key;
                    object obj = comm.ExecuteScalar();
                    if (obj == null || obj == DBNull.Value)
                        continue;
                    ProductStorageDocsP(id_doc, Convert.ToInt32(obj), (int) de.Value, "perso", trans);
                }
            }
            return res;
        }
        public string InsertFromDocs(int id_doc, int id_act, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                        da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card) select @id_doc,id_card from Cards_StorageDocs where id_doc=@id_act";
                        da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                        da.SelectCommand.Parameters.Add("@id_act", SqlDbType.Int).Value = id_act;
                        da.SelectCommand.ExecuteNonQuery();

                        da.SelectCommand.Parameters.Clear();

                        da.SelectCommand.CommandText = "insert into Products_StorageDocs (id_doc,id_prb,cnt_new,cnt_perso,cnt_brak) select @id_doc,id_prb,cnt_new,cnt_perso,cnt_brak from Products_StorageDocs where id_doc=@id_act";
                        da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                        da.SelectCommand.Parameters.Add("@id_act", SqlDbType.Int).Value = id_act;
                        da.SelectCommand.ExecuteNonQuery();

                    }
            return res;
        }
        public int CntProductInStorageDocs(int id_doc, int id_prb, string sost, SqlTransaction trans)
        {
            int cnt = -1;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.CommandText = "select sum(cnt_" + sost + ") as cnt from Products_StorageDocs where id_doc=@id_doc and id_prb=@id_prb";
                    da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prb;
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    object o = da.SelectCommand.ExecuteScalar();
                    cnt = (o == DBNull.Value) ? 0 : Convert.ToInt32(o);
                }
            return cnt;
        }
        public string InsertProductsFromCardsExpertiza(int id_doc, int id_stat, SqlTransaction trans)
        {
            string res = String.Empty;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                DataSet ds = new DataSet();
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;

                ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_stat={0}) and (id_doc={1}) and (id_prop<>10 and id_prop<>1) group by id_prb", id_stat.ToString(), id_doc.ToString()), ref ds, trans);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), "brak", trans);
                ds.Clear();
                ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_stat={0}) and (id_doc={1}) and (id_prop=10 or id_prop=1) group by id_prb", id_stat.ToString(), id_doc.ToString()), ref ds, trans);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), "perso", trans);
            }
            return res;
        }
        public string InsertProductsFromCardsDocReturnToBank(int id_doc, int id_stat, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        DataSet ds = new DataSet();
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;

                        //--------------------------------------------
                        //Object o = null;
                        //int pincountallprop = -1, pincount10prop = -1;
                        //ExecuteScalar(String.Format("select count(*) from Cards_StorageDocs join Cards on Cards.id=Cards_StorageDocs.id_card where ispin!=0 and id_doc={0} and id_prop<>10 and id_stat={1}", id_doc, id_stat), ref o, null);
                        //if (o != null && o!=DBNull.Value) pincountallprop = Convert.ToInt32(o);
                        //ExecuteScalar(String.Format("select count(*) from Cards_StorageDocs join Cards on Cards.id=Cards_StorageDocs.id_card where ispin!=0 and id_doc={0} and id_prop=10 and id_stat={1}", id_doc, id_stat), ref o, null);
                        //if (o != null && o != DBNull.Value) pincount10prop = Convert.ToInt32(o);
                        //--------------------------------------------

                        ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_stat={0}) and (id_doc={1}) and (id_prop<>10) group by id_prb", id_stat.ToString(), id_doc.ToString()), ref ds, trans);
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), "brak", trans);
                        ds.Clear();
                        ExecuteQuery(String.Format("select id_prb,Count(id_prb) as cnt from V_Cards_StorageDocs where (id_stat={0}) and (id_doc={1}) and (id_prop=10) group by id_prb", id_stat.ToString(), id_doc.ToString()), ref ds, trans);
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), "perso", trans);
                    }
            return res;
        }
        public string ProductStorageDocsP(int id_doc, int id_prb, int cnt, string sost, SqlTransaction trans)
        {
            string res = String.Empty;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                        da.SelectCommand.CommandText = "select type from StorageDocs where id=@id_doc";
                        da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                        object obj_tp = da.SelectCommand.ExecuteScalar();
                        da.SelectCommand.CommandText = "select id from Products_StorageDocs where id_prb=@id_prb and id_doc=@id_doc";                        
                        da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prb;
                        object obj = da.SelectCommand.ExecuteScalar();
                        da.SelectCommand.Parameters.Clear();
                        if (obj!=DBNull.Value && Convert.ToInt32(obj) > 0)
                        {
                            da.SelectCommand.CommandText = "update Products_StorageDocs set cnt_" + sost + "=cnt_" + sost + "+@cnt where id_prb=@id_prb and id_doc=@id_doc";
                            da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                            da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                            da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prb;
                            da.SelectCommand.ExecuteNonQuery();
                        }
                        else
                        {
                            da.SelectCommand.CommandText = "insert into Products_StorageDocs (id_doc,id_prb,cnt_" + sost + ") values (@id_doc,@id_prb,@cnt)";
                            da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                            da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prb;
                            da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                            da.SelectCommand.ExecuteNonQuery();
                        }
                        // для карт, возвращающихся из филиалов, прибавляем пины
                        if (Convert.ToInt32(obj_tp) == 9 || Convert.ToInt32(obj_tp) == 10)
                        {
                            da.SelectCommand.Parameters.Clear();
                            da.SelectCommand.CommandText = "select id_prb from V_ProductsBanks_T where id_type=2";
                            obj = da.SelectCommand.ExecuteScalar();
                            int id_pin = (obj == DBNull.Value) ? -1 : Convert.ToInt32(obj);
                            if (id_pin > 0)
                            {
//                                da.SelectCommand.CommandText = "select id from Products_StorageDocs where id_prb=@id_prb and id_doc=@id_doc";
                                //da.SelectCommand.CommandText = "select count(*) from V_Cards_StorageDocs where id_doc=@id_doc and isPin=1";
                                // Пин конверты не правильно были 22.01.2016
                                if (sost == "perso")
                                    da.SelectCommand.CommandText = $"select count(*) from V_Cards_StorageDocs where id_doc={id_doc} and isPin=1 and id_prb={id_prb} and id_prop=10";
                                if (sost == "brak")
                                    da.SelectCommand.CommandText = $"select count(*) from V_Cards_StorageDocs where id_doc={id_doc} and isPin=1 and id_prb={id_prb} and id_prop<>10";
                                obj = da.SelectCommand.ExecuteScalar();
                                int pinCnt = (obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                if (pinCnt > 0)
                                {
                                    da.SelectCommand.Parameters.Clear();
                                    da.SelectCommand.CommandText = "select id from Products_StorageDocs where id_prb=@id_prb and id_doc=@id_doc";
                                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                                    da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_pin;
                                    obj = da.SelectCommand.ExecuteScalar();
                                    da.SelectCommand.Parameters.Clear();
                                    if (Convert.ToInt32(obj) > 0)
                                    {
                                        da.SelectCommand.CommandText = "update Products_StorageDocs set cnt_" + sost + "=cnt_" + sost + "+@cnt where id_prb=@id_prb and id_doc=@id_doc";
                                        da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = pinCnt;
                                        da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                                        da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_pin;
                                        da.SelectCommand.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        da.SelectCommand.CommandText = "insert into Products_StorageDocs (id_doc,id_prb,cnt_" + sost + ") values (@id_doc,@id_prb,@cnt)";
                                        da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                                        da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_pin;
                                        da.SelectCommand.Parameters.Add("@cnt", SqlDbType.Int).Value = pinCnt;
                                        da.SelectCommand.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                    }
            return res;
        }
        protected void gvDocs_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvDocs.PageIndex = e.NewPageIndex;
                Refr(0, false);
            }
        }
        //protected void WriteToLog(string str)
        //{
        //    lock (lockObjectLog)
        //    {
        //        StreamWriter sw = new StreamWriter("c:\\CardPerso\\cardperso.txt", true, System.Text.Encoding.GetEncoding(1251));
        //        sw.WriteLine(String.Format("{1:dd.MM.yy HH:mm:ss}\t{0}", str, DateTime.Now));
        //        sw.Close();
        //    }
        //}

        protected string GetPosition()
        {
            if (userToFilialFilial.Length > 0) return sc.UserPosition(userToFilialFilial);
            string position = null;
            Object obj = gvDocs.DataKeys[gvDocs.SelectedIndex].Values["LoweredUserName"];
            if (obj != null)
            {
                string user = obj.ToString();
                position=sc.UserPosition(user);
            }
            return position;
        }
        protected string GetPositionFio()
        {
            if (userToFilialFilial.Length > 0) return sc.UserFIO(userToFilialFilial);
            string positionFio = null;
            Object obj = gvDocs.DataKeys[gvDocs.SelectedIndex].Values["LoweredUserName"];
            if (obj != null)
            {
                string user = obj.ToString();
                positionFio = sc.UserFIO(user);
            }
            return positionFio;
        }
        protected void bSearchCardNumber_Click(object sender, EventArgs e)
        {
            if (txtCardNumber.Text.Length > 0)
            {
                string pan = txtCardNumber.Text.Replace("-", "");
                string hashpan = FuncClass.GetHashPan(pan);
                string sql = "select cards.id, fio,id_stat,[name] from  cards join status on status.id=cards.id_stat where panhash='" + hashpan + "'";
                SqlCommand cmdSelect = new SqlCommand(sql, conn);
                try
                {
                    lock (Database.lockObjectDB)
                    {
                        using (SqlDataReader dr = cmdSelect.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                dr.Read();
                                string fio = (string) dr["fio"];
                                int id_stat = Convert.ToInt32(dr["id_stat"]);
                                if (id_stat != 8)
                                {
                                    throw new Exception("Карта с номером " + txtCardNumber.Text +
                                                        " находится в состоянии - " + (string) dr["name"]);
                                }

                                txtCardNumberInfo.Text = fio;
                                ClientScript.RegisterClientScriptBlock(GetType(), "okSearch",
                                    "<script type='text/javascript'>" +
                                    "$(document).ready(function(){ clickShowCardNumber('" + txtCardNumber.Text + "','" +
                                    txtCardNumberInfo.Text + "');});</script>");
                            }
                            else throw new Exception("Карта с номером " + txtCardNumber.Text + " не найдена");
                        }
                    }
                }
                catch (Exception em)
                {
                    txtCardNumberInfo.Text = "";
                    ClientScript.RegisterClientScriptBlock(GetType(), "errSearch",
                        "<script type='text/javascript'>" +
                        "$(document).ready(function(){ clickShowCardNumber('" + txtCardNumber.Text + "','" + txtCardNumberInfo.Text + "'); ShowMessage(\"" + em.Message + "\",function f() {$(txtCardNumber).focus();});});</script>");
                }
            }
        }

        protected void bSaveCardNumber_Click(object sender, EventArgs e)
        {
            if (txtCardNumber.Text.Length > 0)
            {
                string pan = txtCardNumber.Text.Replace("-", "");
                string hashpan = FuncClass.GetHashPan(pan);

                try
                {
                    lock (Database.lockObjectDB)
                    {
                        string sql =
                            "select cards.id, fio,id_stat,[name] from  cards join status on status.id=cards.id_stat where panhash='" +
                            hashpan + "'";
                        SqlCommand cmdSelect = new SqlCommand(sql, conn);
                        int id_card = 0;
                        using (SqlDataReader dr = cmdSelect.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                dr.Read();
                                id_card = Convert.ToInt32(dr["id"]);
                                int id_stat = Convert.ToInt32(dr["id_stat"]);
                                if (id_stat != 8)
                                {
                                    throw new Exception("Добавление не возможно. Карта с номером " +
                                                        txtCardNumber.Text + " находится в состоянии - " +
                                                        (string) dr["name"]);
                                }
                            }
                            else throw new Exception("Карта с номером " + txtCardNumber.Text + " не найдена");
                        }

                        // Может уже добавлена?
                        int id_doc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
                        sql = "select id from Cards_StorageDocs where id_doc = " + id_doc.ToString() +
                              " and id_card = " + id_card.ToString();
                        cmdSelect.CommandText = sql;
                        using (SqlDataReader dr = cmdSelect.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                throw new Exception("Карта с номером " + txtCardNumber.Text + " уже добавлена");
                            }
                        }

                        // Добавляем
                        sql = string.Format(
                            "insert into Cards_StorageDocs (id_doc, id_card, comment) values({0},{1},'{2}')", id_doc,
                            id_card, txtClientCompilant.Text);
                        cmdSelect.CommandText = sql;
                        cmdSelect.ExecuteNonQuery();

                        ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                    }
                }
                catch (Exception em)
                {
                    txtCardNumberInfo.Text = "";
                    ClientScript.RegisterClientScriptBlock(GetType(), "errSave",
                        "<script type='text/javascript'>" +
                        "$(document).ready(function(){ ShowMessage(\"" + em.Message + "\");});</script>");
                }
            }
        }

        protected void bCancelCardNumber_Click(object sender, EventArgs e)
        {
            txtCardNumber.Text = "";
            txtCardNumberInfo.Text = "";
            txtClientCompilant.Text = "";
        }



        protected void gvCards_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (gvCards.Rows.Count > 0)
            //if (e.Row.RowType == DataControlRowType.Header)
            {
                if (IsExpertiza() == false)
                {
                    //this.gvCards.Columns[7].HeaderText = "Комментарий";
                    this.gvCards.HeaderRow.Cells[7].Text = "Комментарий";
                }
                else
                {
                    //this.gvCards.Columns[7].HeaderText = "Жалоба клиента";
                    this.gvCards.HeaderRow.Cells[7].Text = "Жалоба клиента";
                }
            }
        }

        protected void bSaveClientCompilant_Click(object sender, EventArgs e)
        {
            try
            {
                string sql = "update Cards_StorageDocs set comment = '" + txtClientCompilantEdit.Text + "' where id = " + gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id"].ToString();
                SqlCommand cmdUpdate = new SqlCommand(sql, conn);
                cmdUpdate.ExecuteNonQuery();
                txtClientCompilantEdit.Text = "";
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.SelectedIndex);
            }
            catch (Exception em)
            {
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.SelectedIndex);
                ClientScript.RegisterClientScriptBlock(GetType(), "errSave",
                        "<script type='text/javascript'>" +
                        "$(document).ready(function(){ ShowMessage(\"" + em.Message + "\");});</script>");

            }
        }

        protected void bCancelClientCompilant_Click(object sender, EventArgs e)
        {
            txtClientCompilantEdit.Text = "";
            ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.SelectedIndex);
        }

        protected void bSaveCardProperty_Click(object sender, EventArgs e)
        {
            try
            {
                string sql = "update Cards_StorageDocs set resultexpertiza = '" + txtClientCompilantResult.Text + "'where id = " + gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id"].ToString();
                sql = sql + "\r\nupdate Cards set id_prop = " + lCardsProperty.SelectedValue + " where id = " + gvCards.DataKeys[Convert.ToInt32(gvCards.SelectedIndex)].Values["id_card"].ToString();
                SqlCommand cmdUpdate = new SqlCommand(sql, conn);
                cmdUpdate.ExecuteNonQuery();
                txtClientCompilantResult.Text = "";
                AddAutoStorageProduct(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]), null);
                ViewProducts(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), 0);
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.SelectedIndex);
                SetButtonCard();
            }
            catch (Exception em)
            {
                ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.SelectedIndex);
                ClientScript.RegisterClientScriptBlock(GetType(), "errSave",
                        "<script type='text/javascript'>" +
                        "$(document).ready(function(){ ShowMessage(\"" + em.Message + "\");});</script>");

            }
        }

        protected void bCancelCardProperty_Click(object sender, EventArgs e)
        {
            txtClientCompilantView.Text = "";
            ViewCards(Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]), gvCards.SelectedIndex);

        }

        protected void bExpertiza_Click(object sender, ImageClickEventArgs e)
        {
            int typebegin = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
            typebegin++;
            string strselect = " set @id_act=(select id from StorageDocs where id_act=@id_act and type=@id_type and priz_gen=1)\r\n";
            if (typebegin > 23)
            {
                typebegin = 23;
                strselect = " set @id_act=(select id from StorageDocs where id=" + Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]).ToString() + " and type=@id_type and priz_gen=1)\r\n";
            }

            string sql = " declare @id_act int\r\n" +
                         " declare @id_type int\r\n" +
                         " declare @id_actold int\r\n" +
                         " declare @id_typeold int\r\n" +
                         " declare @id_card int\r\n" +
                         " set @id_type = " + typebegin.ToString() +
                         " set @id_act = " + Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]).ToString() +
                         " set @id_typeold = @id_type\r\n" +
                         " set @id_actold = @id_act\r\n" +
                         " set @id_card = " + Convert.ToInt32(gvCards.DataKeys[gvCards.SelectedIndex].Values["id_card"]).ToString() + 
                         " bbb:\r\n" +
                         strselect +
                         " if @id_act is not null\r\n" +
                         " begin\r\n" +
                         " if @id_type=23\r\n" +
                         " begin\r\n" +
                         " select @id_act as id_doc, @id_type as id_type, resultexpertiza, Products.name,Branchs.department,TypeStorageDocs.name,StorageDocs.priz_gen\r\n" +
                         " from Cards_StorageDocs\r\n" +
                         " join Cards on Cards.id = Cards_StorageDocs.id_card\r\n" +
                         " join Products_Banks on Products_Banks.id = Cards.id_prb\r\n" +
                         " join Products on Products.id = Products_Banks.id_prod\r\n" +
                         " join StorageDocs on StorageDocs.id = @id_act\r\n" +
                         " join Branchs on Branchs.id = StorageDocs.id_branch\r\n" +
                         " join TypeStorageDocs on TypeStorageDocs.id = @id_type\r\n" +
                         " where id_doc=@id_act and id_card=@id_card\r\n" +
                         " goto aaa\r\n" +
                         " end\r\n" +
                         " set @id_typeold = @id_type" +
                         " set @id_actold = @id_act" +
                         " set @id_type = @id_type + 1\r\n" +
                         " goto bbb\r\n" +
                         " end\r\n" +
                         " select @id_actold as id_doc, @id_typeold as id_type, resultexpertiza, Products.name,Branchs.department,TypeStorageDocs.name as typedoc,StorageDocs.priz_gen\r\n" +
                         " from Cards_StorageDocs\r\n" +
                         " join Cards on Cards.id = Cards_StorageDocs.id_card\r\n" +
                         " join Products_Banks on Products_Banks.id = Cards.id_prb\r\n" +
                         " join Products on Products.id = Products_Banks.id_prod\r\n" +
                         " join StorageDocs on StorageDocs.id = @id_actold\r\n" +
                         " join Branchs on Branchs.id = StorageDocs.id_branch\r\n" +
                         " join TypeStorageDocs on TypeStorageDocs.id = @id_typeold\r\n" +
                         " where id_doc=@id_actold and id_card=@id_card\r\n" +
                         " aaa:";

            try
            {
                lock (Database.lockObjectDB)
                {
                    SqlCommand cmdSelect = new SqlCommand(sql, conn);
                    using (SqlDataReader dr = cmdSelect.ExecuteReader())
                    {

                        string result = null;
                        string resulterr = null;
                        if (dr.HasRows)
                        {
                            dr.Read();

                            if (dr["resultexpertiza"] != DBNull.Value)
                            {
                                if (Convert.ToInt32(dr["priz_gen"]) == 1)
                                {
                                    result = (string) dr["resultexpertiza"];
                                    result = result.Replace("\r\n", "<br>");
                                    result = result.Replace("\t", " ");
                                }
                            }
                            else
                            {
                                if (dr["department"] != DBNull.Value && dr["typedoc"] != DBNull.Value)
                                {
                                    resulterr = string.Format(
                                        "Результат экспертизы не найден\r\nПодразделение: {0}\r\nТип документа: {1}",
                                        (string) dr["department"], (string) dr["typedoc"]);
                                    resulterr = resulterr.Replace("\r\n", "<br>");
                                }
                            }

                            if (result != null)
                            {
                                ClientScript.RegisterClientScriptBlock(GetType(), "okSearch",
                                    "<script type='text/javascript'>" +
                                    "$(document).ready(function(){  ShowMsg('Результат экспертизы', '" + result +
                                    "');});</script>");
                            }
                            else
                            {
                                if (resulterr != null)
                                    ClientScript.RegisterClientScriptBlock(GetType(), "okSearch",
                                        "<script type='text/javascript'>" +
                                        "$(document).ready(function(){  ShowMsg('Результат экспертизы', '" + resulterr +
                                        "');});</script>");
                                else
                                    throw new Exception("Результат экспертизы не найден");
                            }

                        }
                        else throw new Exception("Результат экспертизы не найден");
                    }
                }
            }
            catch (Exception em)
            {
                ClientScript.RegisterClientScriptBlock(GetType(), "errSave",
                       "<script type='text/javascript'>" +
                       "$(document).ready(function(){ ShowMessage(\"" + em.Message + "\");});</script>");

            }

        }

        protected int getIdBranchFromExpertiza()
        {
            int typebegin = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["type"]);
            if (typebegin != (int)TypeDoc.Expertiza) return 0;
            int iddoc = Convert.ToInt32(gvDocs.DataKeys[gvDocs.SelectedIndex].Values["id"]);
            string sql = " declare @id_act int\r\n" +
                         " declare @id_type int\r\n" +
                         " set @id_type = " + typebegin.ToString() +
                         " set @id_act = " + iddoc.ToString() +
                         " ccc:\r\n" +
                         " set @id_act=(select id_act from StorageDocs where id=@id_act and type=@id_type)\r\n" +
                         " if @id_act is not null\r\n" +
                         " begin\r\n" +
                         " set @id_type = @id_type - 1\r\n" +
                         " if @id_type=20\r\n" +
                         "   begin\r\n" +
                         "   select id_branch from StorageDocs where id=@id_act\r\n" +
                         "   goto ddd;\r\n" +
                         "   end\r\n" +
                         "   goto ccc;\r\n" +
                         "   end\r\n" +
                         " else\r\n" +
                         " begin\r\n" +
                         " select 0 as id_branch\r\n" +
                         " end\r\n" +
                         " ddd:\r\n";
            try
            {
                lock (Database.lockObjectDB)
                {
                    SqlCommand cmdSelect = new SqlCommand(sql, conn);
                    using (SqlDataReader dr = cmdSelect.ExecuteReader())
                    {

                        int id_barnch;
                        if (dr.HasRows)
                        {
                            dr.Read();
                            return Convert.ToInt32(dr["id_branch"]);
                        }
                    }
                }
            }
            catch { }
            return 0;
        }

    }
}