using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.Profile;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.IO;
using OstCard.Data;
using System.Data.SqlClient;
using CardPerso.Administration;

namespace CardPerso
{
    public partial class StorDocEdit : System.Web.UI.Page
    {
//        DataSet ds = new DataSet();
//        DataSet ds1 = new DataSet();
        string res = "";
        //int mode = 0; 
//        int id = 0; 
//        int id_type = 0;
        bool AllData = true;
        bool Filial = true;
        bool Perso = true;
        bool FilialDeliver = true;
        int id_deliv = 0; //нужна для работы с рассылкой одним махом
        ServiceClass sc = new ServiceClass();
        SqlConnection conn = null;
        // object lockObject = new object();

        private bool usermk = false;
        int branch_main_filial = 0;
        int current_branch = -1;

        int branchIdMain = 0;
        bool oneclick = false;
       

        protected bool isMainFilial()
        {
            return (branch_main_filial > 0 && branch_main_filial == current_branch);
        }

        
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                return;
            }
            
            lock (Database.lockObjectDB)
            {
                conn = new SqlConnection();
                conn.ConnectionString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
                conn.Open();
                int mode = Convert.ToInt32(Request.QueryString["mode"]);
                //id = Convert.ToInt32(Request.QueryString["id"]);
                //lbIDD.Text = Convert.ToInt32(Request.QueryString["id"]);
                
                AllData = sc.UserAction(User.Identity.Name, Restrictions.AllData);
                Filial = sc.UserAction(User.Identity.Name, Restrictions.Filial);
                Perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
                FilialDeliver = sc.UserAction(User.Identity.Name, Restrictions.FilialDeliver);

                try
                {
                    usermk = Convert.ToBoolean(ConfigurationSettings.AppSettings["EnableRMK"]);
                }
                catch (Exception exception)
                {
                    usermk = false;
                }

                if (mode == 2) Title = "Редактирование";
                else Title = "Добавление";

                current_branch = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(current_branch, sc.UserAction(User.Identity.Name, Restrictions.Perso));
                branchIdMain = BranchStore.getBranchMainFilial(current_branch, false);
                        
                if (!IsPostBack)
                {
                    //bSave.Attributes.Remove("OnClick");
                    //bSave.Attributes.Add("OnClick", "IsVisible=false");
                    //Session["stordocedit_oneclick"] = null;
                    

                    BlockPanel(-1);
                    ZapCombo();
                    if (mode == 2) 
                        ZapFields();
                    else 
                        tbData.Text = String.Format("{0:d}", DateTime.Now);
                }
            }
        }

        private void ZapCombo()
        {
            DataSet ds = new DataSet();
            //ds.Clear();
            
            string sel = "where active=1 ";
            if (Filial && !Perso)
                sel = sel + " and role_tp=1";            
            if (Perso && !Filial)
                sel = sel + " and role_tp=2";
            if (!Perso && !Filial)
                sel = sel + " and id >= 20";
            if (FilialDeliver)
                sel = sel + " or role_tp=3";
            //~!if (branch_main_filial > 0 && branch_main_filial==current_branch)
            //~!    sel = sel + " or id=10 or id=11"; // принудительно добавляем возврат из филиала и уничтожение ценностей
            string query = "select id,name from TypeStorageDocs " + sel + " order by id_sort";
            res = ExecuteQuery(query, ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                #region узкард
                if (FuncClass.ClientType == ClientType.Uzcard)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.PersoCard)
                        dListType.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString())); 
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToFilial)
                        dListType.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
                    continue;
                }
                #endregion

                // Выдача для 4502 убрана по просьбе Рустема 21.03.2016
                //~!if (branch_main_filial > 0 && branch_main_filial == current_branch && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToClient) continue;
                // Передача Филиал-филиал убрана для всего Казанского филиала по просьбе Рустема 11.04.2016, письмо от 08.04.2016
                //~!if (branch_main_filial > 0 && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FilialFilial) continue;
                // При возврате ценностей переименован пункт по просьбе Рустема письмо от 11.04.2016
                //~!if(isMainFilial()==true && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToBank)
                //~!    dListType.Items.Add(new ListItem("Возврат в ОПБК УПЦ", ds.Tables[0].Rows[i]["id"].ToString()));
                //~!else
                //if (branch_main_filial < 1 && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToFilialFilial) continue;    
                if (branchIdMain < 1 && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToFilialFilial)
                    continue;
                //выдачу клиенту через сервис здесь не отображаем
                if (Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToClientService)
                    continue;
                /*
                if (!sc.UserAction(User.Identity.Name, Restrictions.ReceiveToFilialExpertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveToFilialExpertiza) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.SendToExpertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToExpertiza) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ReceiveToExpertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveToExpertiza) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.Expertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.Expertiza) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.SendToPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToPodotchet) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ReceiveToPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveToPodotchet) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.SendToClientFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToClientFromPodotchet) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ReturnFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReturnFromPodotchet) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ReceiveFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveFromPodotchet) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.WriteOfPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.WriteOfPodotchet) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.CardKilling) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int) TypeDoc.KillingCard) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ToBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int) TypeDoc.ToBook124) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.GetBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.GetBook124) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.FromBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromBook124) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ReceiveBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveBook124) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ToGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ToGoz) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.GetGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.GetGoz) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.FromGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromGoz) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.ReceiveGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveGoz) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.GozToPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromGozToPodotchet) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.PodotchetFromGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ToPodotchetFromGoz) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.PodotchetToGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromPodotchetToGoz) continue;
                if (!sc.UserAction(User.Identity.Name, Restrictions.GozFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ToGozFromPodotchet) continue;
                */
                
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ReceiveToFilialExpertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveToFilialExpertiza) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.SendToExpertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToExpertiza) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ReceiveToExpertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveToExpertiza) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.Expertiza) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.Expertiza) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.SendToPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToPodotchet) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ReceiveToPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveToPodotchet) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.SendToClientFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.SendToClientFromPodotchet) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ReturnFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReturnFromPodotchet) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ReceiveFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveFromPodotchet) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.WriteOfPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.WriteOfPodotchet) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.CardKilling) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int) TypeDoc.KillingCard) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ToBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int) TypeDoc.ToBook124) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.GetBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.GetBook124) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.FromBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromBook124) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ReceiveBook124) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveBook124) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ToGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ToGoz) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.GetGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.GetGoz) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.FromGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromGoz) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.ReceiveGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ReceiveGoz) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.GozToPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromGozToPodotchet) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.PodotchetFromGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ToPodotchetFromGoz) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.PodotchetToGoz) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.FromPodotchetToGoz) continue;
                if (!sc.UserAction_v2(User.Identity.Name, Restrictions.GozFromPodotchet) && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == (int)TypeDoc.ToGozFromPodotchet) continue;                 
                

                dListType.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            }
            if (dListType.Items.Count == 0)
                dListType.Items.Add(new ListItem("Нет доступных документов", "-1"));
            if (Convert.ToInt32(Request.QueryString["mode"]) == 1)
            {
                dListType.SelectedIndex = 0;
                dListType_SelectedIndexChanged(this, null);
            }


            ds.Clear();
            res = ExecuteQuery("select id,name from Couriers", ref ds, null);
            dListCr.Items.Add(new ListItem("", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListCr.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            dListCr.SelectedIndex = 0;
            string Passport = sc.UserPassport(User.Identity.Name);
            tbPassportSeries.Text = sc.PassportSeries(Passport);
            tbPassportNumber.Text = sc.PassportNumber(Passport);

            ds.Clear();
            res = ExecuteQuery($"set NUMERIC_ROUNDABORT off set ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, CONCAT_NULL_YIELDS_NULL, QUOTED_IDENTIFIER on select ap.id, ap.secondname + ' ' + ap.firstname + ' ' + case when ap.patronymic is null then '' else ' ' + ap.patronymic end as fio, ap.personnelnumber from AccountablePerson ap inner join aspnet_users u on ap.userId=u.id where dbo.UserMainBranch(u.UserName)={branchIdMain} order by secondname, firstname asc", ref ds, null);
            WriteToLog($"Запрос на подотчет: всего {ds.Tables[0].Rows.Count}");
            dListPerson.Items.Add(new ListItem("Не определено", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                WriteToLog($"{i}: {ds.Tables[0].Rows[i]["fio"].ToString()} ({ ds.Tables[0].Rows[i]["personnelnumber"].ToString()})");
                dListPerson.Items.Add(new ListItem($"{ds.Tables[0].Rows[i]["fio"].ToString()} ({ds.Tables[0].Rows[i]["personnelnumber"].ToString()})", ds.Tables[0].Rows[i]["id"].ToString()));
            }
            RefrPerson();
        }

        private void ZapFields()
        {
            lbInform.Text = "";
            int id = Convert.ToInt32(Request.QueryString["id"]);
            DataSet ds = new DataSet();
            ds.Clear();
            res = ExecuteQuery(String.Format("select * from StorageDocs where id={0}",id), ref ds, null);

            dListType.Enabled = false;

            int id_type = Convert.ToInt32(Convert.ToInt32(ds.Tables[0].Rows[0]["type"]));
            BlockPanel(id_type);
            //добавил RecieveToBank
            if (id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.ReceiveToFilial 
                || id_type == (int)TypeDoc.SendToClient || id_type == (int)TypeDoc.SendToBank 
                || id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.ReturnToFilial
                || id_type == (int)TypeDoc.FilialFilial
                || id_type == (int)TypeDoc.ReceiveToFilialExpertiza || id_type == (int)TypeDoc.SendToExpertiza
                                                     || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza
                                                     || id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet
                                                     || id_type == (int)TypeDoc.SendToClientFromPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet
                                                     || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                                                     || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                                                     || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet
                                                     )
            {
                dListBranch.Enabled = false;
                RefrFilial(true, id_type);
            }

            tbData.Text = String.Format("{0:d}", (ds.Tables[0].Rows[0]["date_doc"]));
            dListType.SelectedIndex = dListType.Items.IndexOf(dListType.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["type"])));
            dListBranch.SelectedIndex = dListBranch.Items.IndexOf(dListBranch.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["id_branch"])));
            dListCr.SelectedIndex = dListCr.Items.IndexOf(dListCr.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["id_courier"])));
            dFilFrom.SelectedIndex = dFilFrom.Items.IndexOf(dFilFrom.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["id_branch"])));
            dFilTo.SelectedIndex =  dFilTo.Items.IndexOf(dFilTo.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["id_courier"])));

            if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.DontReceiveToFilial 
                || id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet 
                || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                || id_type == (int)TypeDoc.ToPodotchetFromGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
            {
                dListAct.Enabled = false;
                RefrAct(true, id_type);
            }
            if (id_type == (int)TypeDoc.FilialFilial)
            {
                pFilFil.Enabled = false;
            }

            dListAct.SelectedIndex = dListAct.Items.IndexOf(dListAct.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["id_act"])));
            tbInvoiceCr.Text = Convert.ToString(ds.Tables[0].Rows[0]["invoice_courier"]);
            tbCourier.Text = Convert.ToString(ds.Tables[0].Rows[0]["courier_name"]);
            tbComment.Text = Convert.ToString(ds.Tables[0].Rows[0]["comment"]);

            int id_doc = Convert.ToInt32(Request.QueryString["id"]);
            ds.Clear();
            res = ExecuteQuery(String.Format("select top 1 * from AccountablePerson_StorageDocs where id_doc={0} order by id desc", id_doc), ref ds, null);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                dListPerson.SelectedIndex = dListPerson.Items.IndexOf(dListPerson.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["id_person"])));
            }
            else dListPerson.SelectedIndex = 0;

        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            bSave.Enabled = false;
            if (!User.Identity.IsAuthenticated)
            {
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                return;
            }            
            lock (Database.lockObjectDB)
            {
                //string er = "sdgdfgdfgdfgfdg";
                //ClientScript.RegisterClientScriptBlock(GetType(), "errRMK", "<script type='text/javascript'>$(document).ready(function(){ ShowMessage('" + er + "');});</script>");
                //if (Session["stordocedit_oneclick"] != null && (bool)Session["stordocedit_oneclick"])
                //{
                //    Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                //    return;
                //}
                Session["stordocedit_oneclick"] = true;
                int mode = Convert.ToInt32(Request.QueryString["mode"]);
                int id_type = Convert.ToInt32(dListType.SelectedItem.Value);
                int id_branch = -1;
                int id_deliver = -1;
                int id_act = -1;

                if ((id_type == (int)TypeDoc.SendToFilial && rbType.SelectedIndex == 0)
                    || id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToFilialPacket
                    || (id_type == (int)TypeDoc.SendToClient)
                    || (id_type == (int)TypeDoc.SendToBank) || (id_type == (int)TypeDoc.ReceiveToBank)
                    || (id_type == (int)TypeDoc.ReturnToFilial) || (id_type == (int)TypeDoc.DontReceiveToFilial)
                    || id_type == (int)TypeDoc.ReceiveToFilialExpertiza || id_type == (int)TypeDoc.SendToExpertiza
                    || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza
                    || id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet
                    || id_type == (int)TypeDoc.SendToClientFromPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet
                    || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                    || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                    || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet
                    )
                {
                    if (dListBranch.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите филиал";
                        dListBranch.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    
                    id_branch = Convert.ToInt32(dListBranch.SelectedItem.Value);

                }

                if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToFilialPacket 
                    || id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.DontReceiveToFilial 
                    || id_type == (int)TypeDoc.SendToExpertiza || id_type == (int)TypeDoc.ReceiveToExpertiza 
                    || id_type == (int)TypeDoc.Expertiza
                    || id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet 
                    || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124
                    || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz
                    )
                {
                    if (dListAct.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите документ";
                        dListAct.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    id_act = Convert.ToInt32(dListAct.SelectedItem.Value);
                }

                if ((id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.SendToFilialFilial) && rbType.SelectedIndex == 1)
                {
                    if (dListDeliver.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите рассылку";
                        dListDeliver.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    id_deliver = Convert.ToInt32(dListDeliver.SelectedItem.Value);
                }
                if (id_type == (int)TypeDoc.FilialFilial)
                {
                    if (dFilFrom.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите филиал отправки";
                        dFilFrom.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    if (dFilTo.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите филиал назначения";
                        dFilTo.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    id_branch = Convert.ToInt32(dFilFrom.SelectedItem.Value);
                }
                if (id_type == (int)TypeDoc.FromWrapping)
                {
                    if (dListAct.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите документ";
                        dListAct.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    id_act = Convert.ToInt32(dListAct.SelectedItem.Value);
                }
                if (id_type == (int)TypeDoc.ToWrapping)
                {
                    if (dListBranch.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите филиал";
                        dListBranch.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    id_branch = Convert.ToInt32(dListBranch.SelectedItem.Value);
                }
                if (id_type == (int)TypeDoc.DontReceiveToFilial)
                {
                    if (tbComment.Text.Trim().Length == 0)
                    {
                        lbInform.Text = "Данный тип документа требует наличие комментария";
                        tbComment.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                }
                if (id_type == (int)TypeDoc.SendToFilialFilial)
                {
                    if (dListBranch.SelectedIndex < 0)
                    {
                        lbInform.Text = "Выберите филиал";
                        dListBranch.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                    id_branch = Convert.ToInt32(dListBranch.SelectedItem.Value);// sc.BranchId(User.Identity.Name);
                }
                if (usermk && tbPassportSeries.Text.Trim().Length < 1)
                {
                    lbInform.Text = "Укажите серию паспорта пользователя";
                    tbPassportSeries.Focus();
                    bSave.Enabled = true;
                    return;
                }
                if (usermk && tbPassportNumber.Text.Trim().Length < 1)
                {
                    lbInform.Text = "Укажите номер паспорта пользователя";
                    tbPassportNumber.Focus();
                    bSave.Enabled = true;
                    return;
                }
                if (usermk)
                    sc.SetUserPassport(User.Identity.Name, tbPassportSeries.Text, tbPassportNumber.Text);

                if (id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.SendToClientFromPodotchet
                    || id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                    || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                    || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
                {
                    if (dListPerson.SelectedIndex <= 0)
                    {
                        lbInform.Text = "Выберите подотчетное лицо";
                        dListPerson.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                }
                if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124
                    || id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz)
                {
                    if (dListBook124Person.SelectedIndex <= 0)
                    {
                        lbInform.Text = "Выберите ответственное лицо";
                        dListBook124Person.Focus();
                        bSave.Enabled = true;
                        return;
                    }
                }
                ////new
                if (mode == 1)
                {
                    if (rbType.SelectedIndex == 1)
                    {
                        DataSet ds1 = new DataSet();
                        ds1.Clear();
                        
                        if (id_type == (int)TypeDoc.SendToFilial)
                        {
                            // 11.09.2018 - в этом запросе выбирает филиалы, в котором есть отправки в нем или в дочернем, но обособленные филиалы учитываются как дочерние
                            // в итоге для родительского филиала документ создается, но карт в нем нет (они должны бы быть в документе для обособленного, но его нет)
                            res = ExecuteQuery(String.Format("select id_branch from V_DeliversBranchs_T where (id_branch in (select id_branchCard from Cards where id_stat = 2) or id_branch in (select id_BranchCard_parent from V_Cards where id_stat=2)) and (id_deliv={0}) order by branch", id_deliver.ToString()), ref ds1, null);
                            #region 19.09.2018 - возвращаю обратно как было раньше, поскольку похоже это было правильно. :((((
                            // делаем выборку, выкидывая карты из родительского для обособленного и добавляя обособленный филиал
                            //res = ExecuteQuery(String.Format(@"select distinct id_branchCard, b.ident_dep from Cards c left join Branchs b on c.id_branchCard = b.id where id_stat=2 and id_branchCard in   
                            //        (select id from branchs where id in (select id_branch from Delivers_Branchs where id_deliv = {0})
                            //        or(isolated = 1 and id_parent in (select id_branch from Delivers_Branchs where id_deliv = {0}))) order by b.ident_dep", id_deliver), ref ds1, null);
                            #endregion
                            for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                                InsertDoc(Convert.ToInt32(ds1.Tables[0].Rows[i][0]), -1, id_type);
                        }
                        if (id_type == (int)TypeDoc.SendToFilialFilial)
                        {
                            SqlCommand count = Database.Conn.CreateCommand();
                            if (id_deliver == 1024)
                            {
                                res = ExecuteQuery(String.Format("select id as id_branch from Branchs where (id in (select id_branchCard from Cards where id_stat = 4 and id_branchCurrent={0}) or id in (select id_BranchCard_parent from V_Cards where id_stat=4 and id_branchCurrent={0}))  and id!={0} order by id", branchIdMain), ref ds1, null);
                            }
                            else
                            {
                                res = ExecuteQuery(String.Format("select id_branch from V_DeliversBranchs_T where (id_branch in (select id_branchCard from Cards where id_stat = 4 and id_branchCurrent={1}) or id_branch in (select id_BranchCard_parent from V_Cards where id_stat=4 and id_branchCurrent={1})) and (id_deliv={0}) order by branch", id_deliver.ToString(), sc.BranchId(User.Identity.Name)), ref ds1, null);
                            }
                            count.CommandText = "select count(*) from Cards where id_stat = 4 and id_BranchCard=@idTo and id_branchCurrent=@idFrom and id_prop = 1 and id not in (select id_card from V_CardsTypeDocs where type = 19 and id_branch=@idTo)";
                            count.Parameters.Add("@idTo", SqlDbType.Int);
                            count.Parameters.Add("@idFrom", SqlDbType.Int).Value = sc.BranchId(User.Identity.Name);
                            if (ds1.Tables.Count > 0)
                            {
                                for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                                {
                                    count.Parameters["@idTo"].Value = Convert.ToInt32(ds1.Tables[0].Rows[i]["id_branch"]);
                                    object obj = count.ExecuteScalar();
                                    if (Convert.ToInt32(obj) > 0)
                                        InsertDoc(Convert.ToInt32(ds1.Tables[0].Rows[i]["id_branch"]), -1, id_type);
                                }
                            }
                        }
                    }
                    else
                        InsertDoc(id_branch, id_act, id_type);
                    Log(sc.UserGuid(User.Identity.Name), String.Format("Добавлен документ '{0}'", dListType.SelectedItem.Text), null);
                }
                //edit
                if (mode == 2)
                {
                    UpdateDoc(id_type);
                    Log(sc.UserGuid(User.Identity.Name), String.Format("Отредактирован документ '{0}' от {1}", dListType.SelectedItem.Text, tbData.Text), null);
                }
                
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
        private void UpdateAccoutablePerson(int id_doc, int id_person)
        {
            SqlCommand sqCom = new SqlCommand();
            sqCom.CommandText =
                "IF NOT EXISTS (SELECT id FROM AccountablePerson_StorageDocs WHERE id_doc=@id_doc)\r\n" +
                "BEGIN\r\n" +
                "INSERT INTO AccountablePerson_StorageDocs (id_doc,id_person) VALUES (@id_doc,@id_person)\r\n" +
                "END\r\n" +
                "ELSE\r\n" +
                "BEGIN\r\n" +
                "UPDATE AccountablePerson_StorageDocs SET id_person=@id_person where id = (SELECT id FROM AccountablePerson_StorageDocs WHERE id_doc=@id_doc)\r\n" +
                "END";
            sqCom.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
            sqCom.Parameters.Add("@id_person", SqlDbType.Int).Value = id_person;
            ExecuteNonQuery(sqCom, null);

        }
        private void InsertDoc(int id_branch,int id_act, int id_type)
        {            
            SqlCommand sqCom = new SqlCommand();
            sqCom.CommandText = "insert into StorageDocs (number_doc,date_doc,type,id_branch,id_courier,invoice_courier,courier_name,priz_gen,comment,id_act,time_doc,user_id,date_time) values(@number_doc,@date_doc,@type,@id_branch,@id_courier,@invoice_courier,@courier_name,@priz_gen,@comment,@id_act, @time_doc, @user_id, @date_time) select @@identity as lid";
            sqCom.Parameters.Add("@number_doc", SqlDbType.Int).Value = GetNumberDoc(Convert.ToDateTime(tbData.Text),null);
            sqCom.Parameters.Add("@type", SqlDbType.Int).Value = id_type;
            sqCom.Parameters.Add("@date_doc", SqlDbType.DateTime).Value = Convert.ToDateTime(tbData.Text);
            sqCom.Parameters.Add("@comment", SqlDbType.VarChar, 100).Value = tbComment.Text;
            sqCom.Parameters.Add("@priz_gen", SqlDbType.Bit).Value = 0;
            DateTime nowDate = DateTime.Now;
            sqCom.Parameters.Add("@time_doc", SqlDbType.VarChar, 8).Value = String.Format("{0:HH:mm:ss}", nowDate);
            sqCom.Parameters.Add("@user_id", SqlDbType.Int).Value = sc.UserId(User.Identity.Name);
            sqCom.Parameters.Add("@date_time", SqlDbType.DateTime).Value = nowDate;
            //получение в филиале разбили на две части, поштучно и пакетно
            if (id_type == (int)TypeDoc.ReceiveToFilial)
            {
                if (rbReceiveType.SelectedValue == "0")
                    id_type = (int)TypeDoc.ReceiveToFilial;
                if (rbReceiveType.SelectedValue == "1")
                    id_type = (int)TypeDoc.ReceiveToFilialPacket;
                sqCom.Parameters["@type"].Value = id_type;
            }

            if (id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.ReceiveToFilial
                || id_type == (int)TypeDoc.ReceiveToFilialPacket
                || (id_type == (int)TypeDoc.SendToClient && id_branch != -1)
                || id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.ReceiveToBank
                || id_type == (int)TypeDoc.ReturnToFilial || id_type == (int)TypeDoc.FilialFilial
                || id_type == (int)TypeDoc.ToWrapping || id_type == (int)TypeDoc.DontReceiveToFilial
                || id_type == (int)TypeDoc.SendToFilialFilial || id_type == (int)TypeDoc.KillingCard
                || ((id_type == (int)TypeDoc.ReceiveToFilialExpertiza ||
                    id_type == (int)TypeDoc.SendToExpertiza ||
                    id_type == (int)TypeDoc.ReceiveToExpertiza ||
                    id_type == (int)TypeDoc.Expertiza ||
                    id_type == (int)TypeDoc.SendToPodotchet ||
                    id_type == (int)TypeDoc.ReceiveToPodotchet ||
                    id_type == (int)TypeDoc.SendToClientFromPodotchet ||
                    id_type == (int)TypeDoc.ReturnFromPodotchet ||
                    id_type == (int)TypeDoc.ReceiveFromPodotchet ||
                    id_type == (int)TypeDoc.WriteOfPodotchet) && id_branch != -1)
                || id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.GetBook124
                || id_type == (int)TypeDoc.ReceiveBook124 || id_type == (int)TypeDoc.FromBook124
                || id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.GetGoz
                || id_type == (int)TypeDoc.FromGoz || id_type == (int)TypeDoc.ReceiveGoz
                || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
            {
                //~!if (id_type == (int)TypeDoc.ReceiveToBank && branch_main_filial > 0)
                //~!{
                    //~!sqCom.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch_main_filial;
                //~!}
                //~!else
                if(id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.KillingCard
                    || id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.GetBook124
                    || id_type == (int)TypeDoc.ReceiveBook124 || id_type == (int)TypeDoc.FromBook124
                    || id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.GetGoz
                    || id_type == (int)TypeDoc.ReceiveGoz || id_type == (int)TypeDoc.FromGoz)
                {
                    sqCom.Parameters.Add("@id_branch", SqlDbType.Int).Value = current_branch;
                }
                else sqCom.Parameters.Add("@id_branch", SqlDbType.Int).Value = id_branch;
            }
            else
                //~!if (id_type == (int)TypeDoc.DeleteBrak && branch_main_filial > 0)
                //~!{
                    //~!sqCom.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch_main_filial;
                //~!}
                //~!else
                    sqCom.Parameters.Add("@id_branch", SqlDbType.Int).Value = DBNull.Value;
            if (Convert.ToInt32(dListCr.SelectedItem.Value) != -1)
                sqCom.Parameters.Add("@id_courier", SqlDbType.Int).Value = Convert.ToInt32(dListCr.SelectedItem.Value);
            else
                sqCom.Parameters.Add("@id_courier", SqlDbType.Int).Value = DBNull.Value;
            sqCom.Parameters.Add("@invoice_courier", SqlDbType.VarChar, 20).Value = tbInvoiceCr.Text;
            sqCom.Parameters.Add("@courier_name", SqlDbType.VarChar, 50).Value = tbCourier.Text;
            sqCom.Parameters.Add("@id_act", SqlDbType.Int).Value = 0;

            if (id_type == (int)TypeDoc.FilialFilial)
            {
                sqCom.Parameters["@id_courier"].Value = Convert.ToInt32(dFilTo.SelectedItem.Value);
                sqCom.Parameters["@id_act"].Value = (cbForKilling.Checked) ? 1 : 0;
            }
            if (id_type == (int)TypeDoc.SendToFilialFilial || id_type == (int)TypeDoc.SendToFilial)
                sqCom.Parameters["@id_act"].Value = sc.BranchId(User.Identity.Name);

            if (id_type == (int)TypeDoc.ReceiveToFilial  || id_type == (int)TypeDoc.ReceiveToFilialPacket
                || id_type == (int)TypeDoc.ReceiveToBank 
                || id_type == (int)TypeDoc.FromWrapping || id_type == (int)TypeDoc.DontReceiveToFilial
                || id_type == (int)TypeDoc.SendToExpertiza || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza
                || id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.SendToClientFromPodotchet
                || id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                || id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124
                || id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz
                || id_type == (int)TypeDoc.ToPodotchetFromGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
                sqCom.Parameters["@id_act"].Value = id_act;
            if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124)
                sqCom.Parameters["@id_act"].Value = sc.UserId(dListBook124Person.SelectedValue.ToString());
            if (id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz)
                sqCom.Parameters["@id_act"].Value = sc.UserId(dListBook124Person.SelectedValue.ToString());

            object obj = null;
            ExecuteScalar(sqCom, ref obj, null);
            int id_doc = Convert.ToInt32(obj);
            WebLog.LogClass.WriteToLog("StorDocEdit.InsertDoc id_doc={0}, id_type={1},id_act = {2}, id_branch = {3}, user = {4}, userbranch = {5}", id_doc, id_type, id_act, id_branch, User.Identity.Name, sc.BranchId(User.Identity.Name));
            WriteToLog(sc.BranchId(User.Identity.Name) + " " + User.Identity.Name + " " +  sc.UserGuid(User.Identity.Name) + "\tДобавлен документ " + id_doc.ToString() + " " + sqCom.Parameters["@number_doc"].Value.ToString() + " " + id_type.ToString());
            sqCom.Parameters.Clear();
            if ((id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.SendToFilialFilial) && (rbType.SelectedIndex == 1))
            {
                if (id_deliv == 0) //нужна для работы с рассылкой одним махом
                    id_deliv = id_doc; //нужна для работы с рассылкой одним махом
                sqCom.CommandText = "update StorageDocs set id_deliv=@id_deliv where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                sqCom.Parameters.Add("@id_deliv", SqlDbType.Int).Value = id_deliv; //нужна для работы с рассылкой одним махом
                ExecuteNonQuery(sqCom, null);
            }
            AddAutoStorageDoc(id_doc, id_type, id_branch,id_act,null);
            if (id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.SendToClientFromPodotchet
                || id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz 
                || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
            {
                UpdateAccoutablePerson(id_doc, Convert.ToInt32(dListPerson.SelectedItem.Value));
            }
        }
        private void UpdateDoc(int id_type)
        {
            int id = Convert.ToInt32(Request.QueryString["id"]);
            SqlCommand sqCom = new SqlCommand();
            sqCom.CommandText = "update StorageDocs set id_courier=@id_courier,invoice_courier=@invoice_courier,courier_name=@courier_name,comment=@comment where id=@id";
            sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
            sqCom.Parameters.Add("@comment", SqlDbType.VarChar, 100).Value = tbComment.Text;

            if (Convert.ToInt32(dListCr.SelectedItem.Value) != -1)
                sqCom.Parameters.Add("@id_courier", SqlDbType.Int).Value = Convert.ToInt32(dListCr.SelectedItem.Value);
            else
                sqCom.Parameters.Add("@id_courier", SqlDbType.Int).Value = DBNull.Value;
            sqCom.Parameters.Add("@invoice_courier", SqlDbType.VarChar, 20).Value = tbInvoiceCr.Text;
            sqCom.Parameters.Add("@courier_name", SqlDbType.VarChar, 50).Value = tbCourier.Text;
            if (id_type == (int)TypeDoc.FilialFilial)
            {
                sqCom.Parameters.Clear();
                sqCom.CommandText = "update StorageDocs set comment=@comment, id_act=@id_act where id=@id";
                sqCom.Parameters.Add("@comment", SqlDbType.VarChar, 100).Value = tbComment.Text;
                sqCom.Parameters.Add("@id_act", SqlDbType.Int).Value = (cbForKilling.Checked) ? 1 : 0;
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
            }
            ExecuteNonQuery(sqCom, null);
            if (id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.SendToClientFromPodotchet
                || id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)

            {
                UpdateAccoutablePerson(id, Convert.ToInt32(dListPerson.SelectedItem.Value));
            }
        }
        protected void dListType_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id_type = Convert.ToInt32(dListType.SelectedItem.Value);
                if (id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.ReceiveToFilial
                    || id_type == (int)TypeDoc.SendToClient || id_type == (int)TypeDoc.SendToBank
                    || id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.ReturnToFilial
                    || id_type == (int)TypeDoc.FilialFilial || id_type == (int)TypeDoc.ToWrapping
                    || id_type == (int)TypeDoc.DontReceiveToFilial || id_type == (int)TypeDoc.SendToFilialFilial
                    || id_type == (int)TypeDoc.ReceiveToFilialExpertiza || id_type == (int)TypeDoc.SendToExpertiza
                    || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza
                    || id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet
                    || id_type == (int)TypeDoc.SendToClientFromPodotchet || id_type == (int)TypeDoc.ReturnFromPodotchet
                    || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet
                    || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                    || id_type == (int)TypeDoc.FromPodotchetToGoz || id_type == (int)TypeDoc.ToGozFromPodotchet
                    )
                    RefrFilial(false, id_type);
                if (id_type == (int)TypeDoc.SendToFilial || id_type == (int)TypeDoc.SendToFilialFilial)
                    RefrDeliver(id_type);
                if (id_type == (int)TypeDoc.FromWrapping)
                    RefrWrapping();
                if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124)
                {
                    #region отправка в Book124
                    dListBook124Person.Items.Clear();
                    dListBook124Person.Items.Add(new ListItem("Не определено", "-1"));
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            int val = (int)Restrictions.GetBook124;
                            if (id_type == (int)TypeDoc.ToBook124)
                                val = (int)Restrictions.GetBook124;
                            if (id_type == (int)TypeDoc.FromBook124)
                                val = (int)Restrictions.ReceiveBook124;
                            comm.CommandText = $"select UserName from V_UserAction where ActionId={val} order by UserName";
                            SqlDataReader dr = comm.ExecuteReader();
                            while (dr.Read())
                            {
                                ProfileBase pb = ProfileBase.Create(dr["UserName"].ToString());
                                UserClass uc = (UserClass)pb.GetPropertyValue("UserData");
                                if (uc.BranchId == current_branch || uc.BranchId == branchIdMain)
                                    dListBook124Person.Items.Add(new ListItem($"{pb.UserName} ({uc.Fio})", pb.UserName));
                            }
                            dr.Close();
                        }
                        conn.Close();
                    }
                    if (dListBook124Person.Items.Count > 0)
                        dListBook124Person.SelectedIndex = 0;
                    #endregion
                }
                if (id_type == (int)TypeDoc.GetBook124)
                {
                    #region получение в Book124
                    dListAct.Items.Clear();                    
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = $"select * from StorageDocs where type={(int)TypeDoc.ToBook124} and priz_gen=1 and id_act={(int)sc.UserId(User.Identity.Name)} and id not in (select id_act from StorageDocs where type={(int)TypeDoc.GetBook124})";
                            SqlDataReader dr = comm.ExecuteReader();
                            while (dr.Read())
                            {
                                dListAct.Items.Add(new ListItem($"Акт {dr["number_doc"].ToString()} от {Convert.ToDateTime(dr["date_doc"]):dd.MM.yy}", dr["id"].ToString()));
                            }
                            dr.Close();
                        }
                        conn.Close();
                    }
                    #endregion
                }
                if (id_type == (int)TypeDoc.ReceiveBook124)
                {
                    dListAct.Items.Clear();
                    #region возврат в хранилище из Book124
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = $"select * from StorageDocs where type={(int)TypeDoc.FromBook124} and priz_gen=1 and id_act={(int)sc.UserId(User.Identity.Name)} and id not in (select id_act from StorageDocs where type={(int)TypeDoc.ReceiveBook124})";
                            SqlDataReader dr = comm.ExecuteReader();
                            while (dr.Read())
                            {
                                dListAct.Items.Add(new ListItem($"Акт {dr["number_doc"].ToString()} от {Convert.ToDateTime(dr["date_doc"]):dd.MM.yy}", dr["id"].ToString()));
                            }
                            dr.Close();
                        }
                        conn.Close();
                    }
                    #endregion
                }
                if (id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz)
                {
                    #region отправка в гоз
                    dListBook124Person.Items.Clear();
                    dListBook124Person.Items.Add(new ListItem("Не определено", "-1"));
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            int val = (int)Restrictions.GetGoz;
                            if (id_type == (int)TypeDoc.ToGoz)
                                val = (int)Restrictions.GetGoz;
                            if (id_type == (int)TypeDoc.FromGoz)
                                val = (int)Restrictions.ReceiveGoz;
                            comm.CommandText = $"select UserName from V_UserAction where ActionId={val} order by UserName";
                            SqlDataReader dr = comm.ExecuteReader();
                            while (dr.Read())
                            {
                                ProfileBase pb = ProfileBase.Create(dr["UserName"].ToString());
                                UserClass uc = (UserClass)pb.GetPropertyValue("UserData");
                                if (uc.BranchId == current_branch || uc.BranchId == branchIdMain)
                                    dListBook124Person.Items.Add(new ListItem($"{pb.UserName} ({uc.Fio})", pb.UserName));
                            }
                            dr.Close();
                        }
                        conn.Close();
                    }
                    if (dListBook124Person.Items.Count > 0)
                        dListBook124Person.SelectedIndex = 0;
                    #endregion
                }
                if (id_type == (int)TypeDoc.GetGoz)
                {
                    #region получение в гоз
                    dListAct.Items.Clear();
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = $"select * from StorageDocs where type={(int)TypeDoc.ToGoz} and priz_gen=1 and id_act={(int)sc.UserId(User.Identity.Name)} and id not in (select id_act from StorageDocs where type={(int)TypeDoc.GetGoz})";
                            SqlDataReader dr = comm.ExecuteReader();
                            while (dr.Read())
                            {
                                dListAct.Items.Add(new ListItem($"Акт {dr["number_doc"].ToString()} от {Convert.ToDateTime(dr["date_doc"]):dd.MM.yy}", dr["id"].ToString()));
                            }
                            dr.Close();
                        }
                        conn.Close();
                    }
                    #endregion
                }
                if (id_type == (int)TypeDoc.ReceiveGoz)
                {
                    dListAct.Items.Clear();
                    #region возврат в хранилище из гоз
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = $"select * from StorageDocs where type={(int)TypeDoc.FromGoz} and priz_gen=1 and id_act={(int)sc.UserId(User.Identity.Name)} and id not in (select id_act from StorageDocs where type={(int)TypeDoc.ReceiveGoz})";
                            SqlDataReader dr = comm.ExecuteReader();
                            while (dr.Read())
                            {
                                dListAct.Items.Add(new ListItem($"Акт {dr["number_doc"].ToString()} от {Convert.ToDateTime(dr["date_doc"]):dd.MM.yy}", dr["id"].ToString()));
                            }
                            dr.Close();
                        }
                        conn.Close();
                    }
                    #endregion
                }

                BlockPanel(id_type);
            }
        }
        protected void dListPerson_OnSelectedIndexChanged(object sender, EventArgs e)
        {
            if (dListType.SelectedItem == null)
                return;
            int id_type = Convert.ToInt32(dListType.SelectedItem.Value);
            if (id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet
                || id_type == (int)TypeDoc.ToPodotchetFromGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
                RefrAct(false, id_type);
        }
        protected void dListBook124Person_OnSelectedIndexChanged(object sender, EventArgs e)
        {
        }
        protected void dListBranch_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id_type = Convert.ToInt32(dListType.SelectedItem.Value);
                if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToBank
                    || id_type == (int)TypeDoc.DontReceiveToFilial)
                    RefrAct(false, id_type);
                if (id_type == (int)TypeDoc.SendToExpertiza || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza)
                    RefrAct(false, id_type);
                if (id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.SendToClientFromPodotchet
                    || id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.ReceiveFromPodotchet 
                    || id_type == (int)TypeDoc.WriteOfPodotchet
                    || id_type == (int)TypeDoc.ToPodotchetFromGoz || id_type == (int)TypeDoc.ToGozFromPodotchet)
                {
                    RefrAct(false, id_type);
                }
            }
        }
        private void RefrPerson()
        {
            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
            {
                conn.Open();
                using (SqlCommand comm = conn.CreateCommand())
                {
                    comm.CommandText = $"select top 1 a.id from AccountablePerson a left join aspnet_users b on a.userid=b.id where b.username='{User.Identity.Name}'";
                    object obj = comm.ExecuteScalar();
                    if (obj != null)
                        dListPerson.SelectedIndex =
                            dListPerson.Items.IndexOf(dListPerson.Items.FindByValue(Convert.ToString(obj)));
                    else
                        dListPerson.SelectedIndex = 0;
                    dListPerson_OnSelectedIndexChanged(null, null);
                }
                conn.Close();
            }

            //DataSet ds = new DataSet();
            //ds.Clear();
            //res = ExecuteQuery(String.Format("select top 1 * from AccountablePerson_StorageDocs where id_doc={0} order by id desc", id_act), ref ds, null);
            //if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    dListPerson.SelectedIndex = dListPerson.Items.IndexOf(dListPerson.Items.FindByValue(Convert.ToString(ds.Tables[0].Rows[0]["id_person"])));
            //}
            //else dListPerson.SelectedIndex = 0;
        }

        protected void dListAct_SelectedIndexChanged(object sender, EventArgs e)
        {
            //RefrPerson(Convert.ToInt32(dListAct.Items[dListAct.SelectedIndex].Value));
        }
        protected void dFilFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        protected void dFilTo_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void BlockPanel(int id_type)
        {
            int mode = Convert.ToInt32(Request.QueryString["mode"]);
            rbType.SelectedIndex = 0;
            pReceiveType.Visible = false;
            pTypeAct.Visible = false;
            pBranch.Visible = false;
            pAct.Visible = false;
            pDeliver.Visible = false;
            pCourier.Visible = false;
            pFilFil.Visible = false;
            pPassport.Visible = usermk;
            pAccountablePerson.Visible = false;
            pBook124.Visible = false;
            pForKilling.Visible = false;


            if (id_type == (int)TypeDoc.SendToFilial)
            {
                if (mode == 1) pTypeAct.Visible = true;
                pBranch.Visible = true;
                pCourier.Visible = true;
            }
            if (id_type == (int)TypeDoc.SendToFilialFilial)
            {
                if (mode == 1) pTypeAct.Visible = true;
                pBranch.Visible = true;
                pCourier.Visible = true;
            }

            if (id_type == (int)TypeDoc.SendToClient || id_type == (int)TypeDoc.SendToBank || id_type == (int)TypeDoc.ReceiveToBank 
                || id_type == (int)TypeDoc.ReturnToFilial
                || id_type == (int)TypeDoc.ReceiveToFilialExpertiza || id_type == (int)TypeDoc.SendToExpertiza
                || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza
                || id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.ReceiveToPodotchet
                || id_type == (int)TypeDoc.FromGozToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz
                )

            {
                pBranch.Visible = true;
            }

            if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.DontReceiveToFilial)
            {
                pBranch.Visible = true;
                pAct.Visible = true;
            }
            if (id_type == (int)TypeDoc.ReceiveToFilial)
            {
                pReceiveType.Visible = true;
            }
            if (id_type == (int)TypeDoc.FilialFilial)
            {
                pFilFil.Visible = true;
                pForKilling.Visible = true;
            }
            if (id_type == (int)TypeDoc.ToWrapping)
                pBranch.Visible = true;
            if (id_type == (int)TypeDoc.FromWrapping)
                pAct.Visible = true;
            if (id_type == (int)TypeDoc.SendToExpertiza || id_type == (int)TypeDoc.SendToExpertiza
                                                        || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza)
            {
                pBranch.Visible = true;
                pAct.Visible = true;
            }
            if (id_type == (int) TypeDoc.SendToPodotchet)
            {
                pBranch.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = true;
            }
            if (id_type == (int) TypeDoc.SendToClientFromPodotchet)
            {
                //pAct.Visible = true;
                pPassport.Visible = false;
                pBranch.Visible = true;
                pAccountablePerson.Visible = true;
                //pAccountablePerson.Enabled = false;
            }
            if (id_type == (int) TypeDoc.ReturnFromPodotchet)
            {
                pBranch.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = false;
            }
            if (id_type == (int)TypeDoc.ReceiveToPodotchet )
            {
                pBranch.Visible = true;
                pAct.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = false;
            }
            if (id_type == (int)TypeDoc.ReceiveFromPodotchet)
            {
                pBranch.Visible = true;
                //pAct.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = true;
            }
            if (id_type == (int)TypeDoc.FromGozToPodotchet)
            {
                pBranch.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = true;
            }            
            if (id_type == (int)TypeDoc.FromPodotchetToGoz)
            {
                pBranch.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = false;
            }
            if (id_type == (int)TypeDoc.ToPodotchetFromGoz)
            {
                pBranch.Visible = true;
                pAct.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = false;
            }
            if (id_type == (int)TypeDoc.ToGozFromPodotchet)
            {
                pBranch.Visible = true;
                //pAct.Visible = true;
                pAccountablePerson.Visible = true;
                pAccountablePerson.Enabled = true;
            }

            if (id_type == (int)TypeDoc.ToBook124 || id_type == (int)TypeDoc.FromBook124)
            {
                pBook124.Visible = true;
            }
            if (id_type == (int)TypeDoc.GetBook124 || id_type == (int)TypeDoc.ReceiveBook124)
            {
                pAct.Visible = true;
            }
            if (id_type == (int)TypeDoc.ToGoz || id_type == (int)TypeDoc.FromGoz)
            {
                pBook124.Visible = true;
            }
            if (id_type == (int)TypeDoc.GetGoz || id_type == (int)TypeDoc.ReceiveGoz)
            {
                pAct.Visible = true;
            }
        }
        private void RefrFilial(bool show_all, int id_type)
        {
            dListBranch.Items.Clear();
            dFilFrom.Items.Clear();
            dFilTo.Items.Clear();
            DataSet ds1 = new DataSet();
            ds1.Clear();

            string dop = "";
            string dop1 = "";
            int fil = sc.BranchId(User.Identity.Name);
            if (!AllData)
            {
                dop = " and id_branch=" + sc.BranchId(User.Identity.Name).ToString();
                dop1 = " and id_courier=" + sc.BranchId(User.Identity.Name).ToString();
            }
            //if (FilialDeliver)
            //{
            //    string kz = ConfigurationSettings.AppSettings["Kazan"];
            //    SqlCommand sel = new SqlCommand();
            //    sel.Connection = Database.Conn;
            //    sel.CommandText = "select id from Branchs where ident_dep='"+kz+"'";
            //    object obj = null;
            //    ExecuteScalar(sel,ref obj, null);
            //    if (obj != null)
            //        dop_kazan = String.Format(" and (id_parent={0})", (int)obj);
            //}

            if (show_all && id_type != (int)TypeDoc.FilialFilial)
                res = ExecuteQuery("select id,department from Branchs order by department", ref ds1, null);
            else
            {
                if (id_type == (int)TypeDoc.SendToFilial)
                    //~!res = ExecuteQuery(String.Format("select id, department from Branchs where ((id in (select id_branchCard_parent from V_Cards where id_stat = 2 and id_prop=1)) or (id in (select id_BranchCard from Cards where id_stat = 2 and id_prop=1))) order by department"), ref ds1, null);
                    res = ExecuteQuery(String.Format("select id, department from Branchs where id in (select id_BranchCard from Cards where id_stat = 2 and id_prop=1) order by department"), ref ds1, null);
                if (id_type == (int)TypeDoc.SendToFilialFilial)
                    res = ExecuteQuery(String.Format("select id, department from Branchs where id in (select id_branchCard from Cards where id_stat = 4 and id_prop=1 and id_branchCurrent={0} and id_branchCard<>{0}) order by department", fil), ref ds1, null);
                if (id_type == (int)TypeDoc.ReceiveToFilial)
                    //res = ExecuteQuery("select id, department from Branchs where (id in (select id_branchCard_parent from V_Cards where id_stat = 3)) or (id in (select id_BranchCard from Cards where id_stat = 3))", ref ds1, null);
                    res = ExecuteQuery(String.Format("select distinct id_branch as id,branch as department from V_StorageDocs where ((type = 5 or type = 19) and priz_gen = 1) and (id not in (select id_act from StorageDocs)) {0} order by department", dop), ref ds1, null);
                if (id_type == (int)TypeDoc.DontReceiveToFilial)
                    //res = ExecuteQuery("select id, department from Branchs where (id in (select id_branchCard_parent from V_Cards where id_stat = 3)) or (id in (select id_BranchCard from Cards where id_stat = 3))", ref ds1, null);
                    //res = ExecuteQuery(String.Format("select distinct id_branch as id,branch as department from V_StorageDocs where (type = 5 and priz_gen = 1) and (id not in (select id_act from StorageDocs)) {0} order by department", dop), ref ds1, null);
                    res = ExecuteQuery(String.Format("select distinct id_branch as id,branch as department from V_StorageDocs where (type in(5,6) and priz_gen = 1) and (id not in (select id_act from StorageDocs)) {0} order by department", dop), ref ds1, null);

                if (id_type == (int)TypeDoc.SendToBank)
                {
                    //res = ExecuteQuery("select id, department from Branchs where (id in (select id_branchCard_parent from V_Cards where id_stat = 4)) or (id in (select id_BranchCard from Cards where id_stat = 4))", ref ds1, null);
                    //res = ExecuteQuery(String.Format("select distinct id_branch as id,branch as department from V_StorageDocs where (type = 5) and (priz_gen = 1) {0} order by department",dop), ref ds1, null);
                    //~!if (isMainFilial() == false)
                    //~!res = ExecuteQuery(String.Format("select distinct id_branch as id,branch as department from V_StorageDocs where (type in(5,6) and (priz_gen = 1) {0}) order by department", dop), ref ds1, null);
                    //~!else
                    res = ExecuteQuery(String.Format("select distinct id_branch as id,branch as department from V_StorageDocs where (type in(5,6) and (priz_gen = 1) {0}) order by department", dop), ref ds1, null);

                }
                if (id_type == (int)TypeDoc.ReceiveToBank)
                {
                    ////res = ExecuteQuery("select id, department from Branchs where (id in (select id_branchCard_parent from V_Cards where id_stat = 5)) or (id in (select id_BranchCard from Cards where id_stat = 5))", ref ds1, null);
                    //~!//
                    res = ExecuteQuery("select distinct id_branch as id,branch as department from V_StorageDocs where (type = 9 or type =18) and (priz_gen = 1) and (id not in (select id_act from StorageDocs)) order by department", ref ds1, null);
                    //~!if (branch_main_filial > 0)
                    //~!{
                    //~!res = ExecuteQuery(String.Format("select distinct id_branch as id,branch as department from V_StorageDocs where (type = 9 or type =18) and (priz_gen = 1) and (id not in (select id_act from StorageDocs)) and id_branch in (select id from Branchs where id_parent={0} or id={0}) order by department", branch_main_filial), ref ds1, null);
                    //~!}
                    //~!if (branch_main_filial==0)
                    //~!{
                    //~!    //!!!! переделать запрос, когда Казанский будет не один (убрать 106)
                    //~!    //res = ExecuteQuery("select distinct id_branch as id,branch as department from V_StorageDocs where (type = 9 or type =18) and (priz_gen = 1) and (id not in (select id_act from StorageDocs)) and id_branch not in (select id from Branchs where (id_parent=0 or id_parent is null or id_parent>0) and (type!=18 and (id=106 or id_parent=106))) order by department", ref ds1, null);                    
                    //~!    //Чтобы можно было получить карты из Казанского с неверным филиалом выдачи 11.04.2016 Рустем
                    //~!    res = ExecuteQuery("select distinct id_branch as id,branch as department from V_StorageDocs" + 
                    //~!                        " join Cards_StorageDocs on Cards_StorageDocs.id_doc=V_StorageDocs.id" +
                    //~!                        " join Cards on Cards.id=Cards_StorageDocs.id_card where (type = 9 or type =18) and" + 
                    //~!                        " (priz_gen = 1) and (V_StorageDocs.id not in (select id_act from StorageDocs)) and" + 
                    //~!                        " id_branch not in (select id from Branchs where ([type]!=18 and (id=106 or id_parent=106)))" +
                    //~!                        " or (type=9 and id_branch=106 and Cards.id_prop=10) order by department", ref ds1, null);                    
                    //~!}
                }
                if (id_type == (int)TypeDoc.ToWrapping)
                    res = ExecuteQuery("select id, department from Branchs where (id in (select id_branchCard_parent from V_Cards where id_stat=9 and id_prop=1)) or (id in (select id_BranchCard from Cards where id_stat=9 and id_prop=1)) order by department", ref ds1, null);
                if (id_type == (int)TypeDoc.SendToClient)
                {
                    if (!AllData)
                        dop = " and id_branchCard=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select distinct id_BranchCard as id,DepBranchCard as department from V_Cards where (id_stat=4 or id_stat=19) and (id_prop=1) and (DepBranchCard is not null) and (DateClient is null) {0}", dop), ref ds1, null);
                }
                if (id_type == (int)TypeDoc.ReturnToFilial)
                {
                    if (!AllData)
                        dop = " and id_branchCard=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select distinct id_BranchCard as id,DepBranchCard as department from V_Cards where (id_stat=8) and (DepBranchCard is not null) {0}", dop), ref ds1, null);
                }
                if (id_type == (int)TypeDoc.FilialFilial)
                {
                    if (!AllData)
                        dop = " and id_branchCard=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select distinct id_BranchCard as id,DepBranchCard as department from V_Cards where (id_stat=4) and (DateClient is null) {0}", dop), ref ds1, null);
                    for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                        dFilFrom.Items.Add(new ListItem(ds1.Tables[0].Rows[i]["department"].ToString(), ds1.Tables[0].Rows[i]["id"].ToString()));
                    ds1.Clear();
                    res = ExecuteQuery("select id,department from Branchs order by department", ref ds1, null);
                    for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                        dFilTo.Items.Add(new ListItem(ds1.Tables[0].Rows[i]["department"].ToString(), ds1.Tables[0].Rows[i]["id"].ToString()));
                }
                if (id_type == (int)TypeDoc.SendToFilialFilial)
                {

                }
                if (id_type == (int)TypeDoc.ReceiveToFilialExpertiza)
                {
                    if (!AllData)
                        dop = "where id=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select id, department from Branchs {0}", dop), ref ds1, null);

                }

                if (id_type == (int)TypeDoc.SendToExpertiza)
                {
                    if (!AllData)
                        dop = "where Branchs.id=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={1} and StorageDocs.priz_gen!=0 {0}", dop, (int)TypeDoc.ReceiveToFilialExpertiza), ref ds1, null);

                }
                if (id_type == (int)TypeDoc.ReceiveToExpertiza)
                {
                    res = ExecuteQuery(String.Format("select Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={0} and StorageDocs.priz_gen!=0", (int)TypeDoc.SendToExpertiza), ref ds1, null);
                }
                if (id_type == (int)TypeDoc.Expertiza)
                {
                    if (!AllData)
                        dop = "where Branchs.id=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={1} and StorageDocs.priz_gen!=0 {0}", dop, (int)TypeDoc.ReceiveToExpertiza), ref ds1, null);

                }

                if (id_type == (int)TypeDoc.SendToPodotchet || id_type == (int)TypeDoc.FromGozToPodotchet)
                {
                    if (!AllData)
                        dop = "where id=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select distinct id, department from Branchs {0}", dop), ref ds1, null);
                }

                if (id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz)
                {
                    if (!AllData)
                        dop = "where Branchs.id=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select distinct Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={1} and StorageDocs.priz_gen!=0 {0}", dop, (int)TypeDoc.SendToPodotchet), ref ds1, null);
                }

                if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
                {
                    if (!AllData)
                        dop = "where Branchs.id=" + sc.BranchId(User.Identity.Name).ToString();
                    //res = ExecuteQuery(String.Format("select distinct Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={1} and StorageDocs.priz_gen!=0 {0}", dop, (int)TypeDoc.ReceiveToPodotchet), ref ds1, null);
                    res = ExecuteQuery(String.Format("select distinct Branchs.id, department from Branchs {0}", dop, (int)TypeDoc.ReceiveToPodotchet), ref ds1, null);
                }

                if (id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.FromPodotchetToGoz)
                {
                    if (!AllData)
                        dop = "where Branchs.id=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select distinct Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={1} and StorageDocs.priz_gen!=0 {0}", dop, (int)TypeDoc.ReceiveToPodotchet), ref ds1, null);
                }
                if (id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.ToGozFromPodotchet)
                {
                    if (!AllData)
                        dop = "where Branchs.id=" + sc.BranchId(User.Identity.Name).ToString();
                    //res = ExecuteQuery(String.Format("select distinct Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={1} and StorageDocs.priz_gen!=0 {0}", dop, (int)TypeDoc.ReturnFromPodotchet), ref ds1, null);
                    //11.12.2019 теперь не смотрим есть ли возврат от подотчетного. Сразу принимаем от него
                    res = ExecuteQuery(String.Format("select distinct id, department from Branchs {0}", dop), ref ds1, null);
                }
                if (id_type == (int)TypeDoc.WriteOfPodotchet)
                {
                    if (!AllData)
                        dop = "where Branchs.id=" + sc.BranchId(User.Identity.Name).ToString();
                    res = ExecuteQuery(String.Format("select distinct Branchs.id, department from Branchs join StorageDocs on StorageDocs.id_branch=Branchs.id and StorageDocs.type={1} and StorageDocs.priz_gen!=0 {0}", dop, (int)TypeDoc.SendToClientFromPodotchet), ref ds1, null);
                }
            }
            if (sc.UserAction(User.Identity.Name, Restrictions.Perso))
            {
                if (id_type == (int)TypeDoc.SendToClient && ds1.Tables[0].Rows.Count > 0)
                {
                    dListBranch.Items.Add(new ListItem("Все", "-1"));
                }
                if (id_type == (int)TypeDoc.ToWrapping && ds1.Tables[0].Rows.Count > 0)
                {
                    dListBranch.Items.Add(new ListItem("Все", "-1"));
                }
            }

            for (int i = 0; ds1.Tables.Count>0 && i < ds1.Tables[0].Rows.Count; i++)
            {
               dListBranch.Items.Add(new ListItem(ds1.Tables[0].Rows[i]["department"].ToString(), ds1.Tables[0].Rows[i]["id"].ToString()));
            }
            if (id_type == (int)TypeDoc.ReceiveToFilial)
            {
                ds1.Clear();
                res = ExecuteQuery(String.Format("select distinct id_courier as id from V_StorageDocs where (type = 13 and priz_gen = 1) and (id not in (select id_act from StorageDocs)) {0}", dop1), ref ds1, null);
                for (int i = 0; i< ds1.Tables[0].Rows.Count; i++)
                {
                    object obj = null;
                    SqlCommand comm = conn.CreateCommand();
                    comm.CommandText = String.Format("select department from Branchs where id=" + ds1.Tables[0].Rows[i][0].ToString());
                    res = ExecuteScalar(comm, ref obj, null);
                    dListBranch.Items.Add(new ListItem(obj.ToString(), ds1.Tables[0].Rows[i]["id"].ToString()));
                }
            }

            if (dListBranch.Items.Count > 0)
            {
                dListBranch.SelectedIndex = 0;
                if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.ReceiveToBank || id_type == (int)TypeDoc.DontReceiveToFilial) 
                    RefrAct(false, id_type);
                if (id_type == (int)TypeDoc.SendToExpertiza || id_type == (int)TypeDoc.ReceiveToExpertiza || id_type == (int)TypeDoc.Expertiza)
                    RefrAct(false, id_type);
                //11.12.2019 для подотчетников убрали все акты
                //if (id_type == (int)TypeDoc.ReceiveToPodotchet 
                //    || id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.WriteOfPodotchet)
                //    RefrAct(false, id_type);
            }
        }
        private void RefrAct(bool show_all, int id_type)
        {
            int id_branch = 0;
            if (dListBranch.SelectedItem != null)
                id_branch = Convert.ToInt32(dListBranch.SelectedItem.Value);
            string sel = "",sel1="",sel2="",sel3 ="";
            
            if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.DontReceiveToFilial)
                sel = "((s.type=5 or s.type=19) and s.priz_gen=1)"; //отправка в филиала из центра или из другого филиала
            if (id_type == (int)TypeDoc.ReceiveToBank)
            {
                sel = "((s.type=9 or s.type=18) and s.priz_gen=1)"; //возврат ценностей или отказ от приема в филиале
                // 19.10.2015 добавление актов возврата из подчиненных филиалов
                //~!if (isMainFilial()) // Казанский или наподобие ему
                //~!{
                //~!    //sel1 = " or (id_branch in (select id from branchs where id_parent=" + current_branch.ToString() + ")) ";
                //~!    //sel2 = " ,(select ident_dep from branchs b where b.id=id_branch) as [ident_dep]";
                //~!    //sel = "(priz_gen=1 and (type=9 or (type=18 and id_branch!=" + current_branch + ")))";
                //~!}
            }
            if (id_type == (int)TypeDoc.SendToExpertiza)
            {
                sel = "(s.type=20 and s.priz_gen=1)";
            }

            if (id_type == (int)TypeDoc.ReceiveToExpertiza)
            {
                sel = "(s.type=21 and s.priz_gen=1)";
            }
            if (id_type == (int)TypeDoc.Expertiza)
            {
                sel = "(s.type=22 and s.priz_gen=1)";
            }

            if (id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz)
            {
                int personid = -1;
                if (dListPerson.SelectedItem != null)
                {
                    personid = Convert.ToInt32(dListPerson.SelectedItem.Value);
                    sel3 = " left join AccountablePerson_StorageDocs asd on s.id=asd.id_doc ";
                    sel = $"(s.type=24 and s.priz_gen=1 and asd.id_person={personid})";
                }
                else
                    sel = $"(s.type=24 and s.priz_gen=1)";
            }
            if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
            {
                sel = $"(s.type=25 and s.priz_gen=1 and id_person={Session["CurrentUserId"]})";
            }
            if (id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.FromPodotchetToGoz)
            {
                sel = "(s.type=25 and s.priz_gen=1)";
            }
            if (id_type == (int)TypeDoc.ReceiveFromPodotchet || id_type == (int)TypeDoc.ToGozFromPodotchet)
            {
                int personid = -1;
                if (dListPerson.SelectedItem != null)
                {
                    personid = Convert.ToInt32(dListPerson.SelectedItem.Value);
                    sel3 =" left join AccountablePerson_StorageDocs asd on s.id=asd.id_doc ";
                    sel = $"(s.type=27 and s.priz_gen=1 and asd.id_person={personid})";
                }
                else
                    sel = $"(s.type=27 and s.priz_gen=1)";
            }

            if (id_type == (int)TypeDoc.WriteOfPodotchet)
            {
                sel = "(s.type=26 and s.priz_gen=1)";
            }
            dListAct.Items.Clear();
            DataSet ds1 = new DataSet();
            ds1.Clear();
            if (show_all)
                //res = ExecuteQuery(String.Format("select s.id, s.date_doc {3} from StorageDocs w {4} where ({0}) and (s.id_branch={1}{2})", sel, id_branch.ToString(), sel1, sel2, sel3), ref ds1, null);
                //07.02.2020 почему то alias таблицы w, а не s!!! Заметил в Казани
                res = ExecuteQuery(String.Format("select s.id, s.date_doc {3} from StorageDocs s {4} where ({0}) and (s.id_branch={1}{2})", sel, id_branch.ToString(), sel1, sel2, sel3), ref ds1, null);
            else
            {
                //res = ExecuteQuery(
                //    String.Format(
                //        "select id, date_doc {3} from StorageDocs where ({0}) and (id_branch={1}{2}) and (id not in (select id_act from StorageDocs))",
                //        sel, id_branch.ToString(), sel1, sel2), ref ds1, null);
                {
                    if (id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.FromPodotchetToGoz)
                        res = ExecuteQuery(String.Format("select s.id, s.date_doc {3} from StorageDocs s {4} where ({0}) and (s.id_branch={1}{2}) and (s.id not in (select id_act from StorageDocs where [type]!=26))", sel, id_branch.ToString(), sel1, sel2, sel3), ref ds1, null);
                    else
                        res = ExecuteQuery(String.Format("select s.id, s.date_doc {3} from StorageDocs s {4} where ({0}) and (s.id_branch={1}{2}) and (s.id not in (select id_act from StorageDocs))", sel, id_branch.ToString(), sel1, sel2, sel3), ref ds1, null);
                }
            }
            for (int i = 0; ds1.Tables.Count > 0 && i < ds1.Tables[0].Rows.Count; i++)
            {
                if(sel2.Length<1) dListAct.Items.Add(new ListItem(String.Format("Акт от {0:d}", (ds1.Tables[0].Rows[i]["date_doc"])), ds1.Tables[0].Rows[i]["id"].ToString()));
                // Добавлен вывод откуда пришел возврат
                else dListAct.Items.Add(new ListItem(String.Format("Акт от {0:d} из {1}", (ds1.Tables[0].Rows[i]["date_doc"]), (ds1.Tables[0].Rows[i]["ident_dep"])), ds1.Tables[0].Rows[i]["id"].ToString()));
            }
            if (id_type == (int)TypeDoc.ReceiveToFilial || id_type == (int)TypeDoc.DontReceiveToFilial)
            {
                ds1.Clear();
                if (show_all)
                    res = ExecuteQuery(String.Format("select s.id, s.date_doc from StorageDocs s where (s.type=13) and (s.id_courier={0})", id_branch),ref ds1, null);
                else
                    res = ExecuteQuery(String.Format("select s.id, s.date_doc from StorageDocs s where (s.type=13) and (s.id_courier={0}) and (s.id not in (select id_act from StorageDocs))", id_branch.ToString()), ref ds1, null);
                for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                    dListAct.Items.Add(new ListItem(String.Format("Акт от {0:d}", (ds1.Tables[0].Rows[i]["date_doc"])), ds1.Tables[0].Rows[i]["id"].ToString()));
            }

            if (dListAct.Items.Count > 0)
            {
                dListAct.SelectedIndex = 0;
                //RefrPerson(Convert.ToInt32(dListAct.Items[dListAct.SelectedIndex].Value));
            }
        }
        private void RefrWrapping()
        {
            dListAct.Items.Clear();            
            DataSet ds1 = new DataSet();
            //здесь было так, что можно делать только один возврат с упаковки
            //res = ExecuteQuery(String.Format("select id, number_doc, date_doc from StorageDocs where type=16 and (id not in (select id_act from StorageDocs))"), ref ds1, null);
            //теперь делаем возможность возврата частями
            res = ExecuteQuery(String.Format("select distinct id_doc as id, number_doc, date_doc from V_CardsTypeDocs where type=16 and id_card in (select id from cards where id_stat=10)" +
                " and id_doc not in (select id_act from StorageDocs where id_act=id_doc and type=17)"), // По Стасу от 27.07.2015, чтобы не было двух одинаковых возвратов
                ref ds1, null);
            for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                dListAct.Items.Add(new ListItem(String.Format("Акт {1} от {0:d}", ds1.Tables[0].Rows[i]["date_doc"], ds1.Tables[0].Rows[i]["number_doc"]), ds1.Tables[0].Rows[i]["id"].ToString()));
            if (dListAct.Items.Count > 0)
                dListAct.SelectedIndex = 0;
           
        }
        private void RefrDeliver(int id_type)
        {
            dListDeliver.Items.Clear();
            DataSet ds1 = new DataSet();
            ds1.Clear();
            string sel = "select distinct id_deliv, deliver from V_DeliversBranchs_T where (id_branch in (select id_branchCard from Cards where id_stat = 2)) and type = 1";
            if (id_type == (int)TypeDoc.SendToFilialFilial)// && FilialDeliver)
            {
                ////sel = String.Format("select distinct id_deliv, deliver from V_DeliversBranchs_T where (id_branch in (select id_branchCard from Cards where id_stat = 4 and id_branchCurrent={0})) and type = 2", sc.BranchId(User.Identity.Name));
                //sel = String.Format("select distinct d.id as id_deliv, d.name as deliver from Delivers d join Delivers_Branchs db on d.id=db.id_deliv join Branchs b on b.id=db.id_branch join Cards c on c.id_branchCard=b.id and c.id_stat=4 and c.id_branchCurrent={0} where d.type=2", sc.BranchId(User.Identity.Name));
                if (branchIdMain > 0 && branchIdMain == current_branch) // Головной офис
                {
                    sel = "select 1024 as id_deliv, department as deliver from branchs where id=" + branchIdMain.ToString();
                }
                else
                {
                    sel = "select distinct id_deliv, deliver from V_DeliversBranchs_T where type = 5000";
                }

            }
            

            res = ExecuteQuery(sel, ref ds1, null);
                   
            dListDeliver.DataSource = ds1.Tables[0];
            dListDeliver.DataTextField = "deliver";
            dListDeliver.DataValueField = "id_deliv";
            dListDeliver.DataBind();

            if (dListDeliver.Items.Count > 0)
                dListDeliver.SelectedIndex = 0;

        }

        protected void rbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rbType.SelectedIndex == 0)
            {
                pBranch.Visible = true;
                pDeliver.Visible = false;
            }
            else
            {
                pDeliver.Visible = true;
                pBranch.Visible = false;
//                pDeliver.Enabled = !FilialDeliver;
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
                    dataAdapter.Fill(dataSet);
                }
            }
            catch (Exception e)
            {
                WriteToLog($"query={query}, err={e.Message}");
            }
            return "";
        }
        public string Log(string UserGuid, string action, SqlTransaction trans)
        {
            string res = "";
                    SqlCommand comm = conn.CreateCommand();
                    if (trans != null)
                        comm.Transaction = trans;
                    comm.CommandTimeout = conn.ConnectionTimeout;
                    comm.CommandText = "insert into LogAction (UserId, ActionDate, Description) values (@uid, @adate, @desc)";
                    comm.Parameters.Add("@uid", SqlDbType.UniqueIdentifier).Value = new Guid(UserGuid);
                    comm.Parameters.Add("@adate", SqlDbType.DateTime).Value = DateTime.Now;
                    comm.Parameters.Add("@desc", SqlDbType.NVarChar, 500).Value = action;
                    comm.ExecuteNonQuery();
            return res;
        }
        public int GetNumberDoc(DateTime dt, SqlTransaction trans)
        {
            int NumberDoc = 0;
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        da.SelectCommand = new SqlCommand("", conn);
                        if (trans != null)
                            da.SelectCommand.Transaction = trans;
                        da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                        // нумерация в рамках одного дня
                        //                        da.SelectCommand.CommandText = "select Max(number_doc) from StorageDocs where date_doc=@date_doc";
                        //                        da.SelectCommand.Parameters.Add("@date_doc", SqlDbType.DateTime).Value = dt;
                        // нумерация общая
                        da.SelectCommand.CommandText = "select Max(number_doc) from StorageDocs";
                        object obj = da.SelectCommand.ExecuteScalar();
                        NumberDoc = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                    }
            return NumberDoc + 1;
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
                WebLog.LogClass.WriteToLogErr("StorDocEdit.ExecuteScalar: commandText={0}, err={1}", comm.CommandText, e.Message);
            }
            return "";
        }
        public string ExecuteNonQuery(SqlCommand comm, SqlTransaction transaction)
        {
            string res = String.Empty;
            try
            {
             
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                comm.Connection = conn;
                comm.CommandTimeout = conn.ConnectionTimeout;
                if (transaction != null)
                    comm.Transaction = transaction;
                comm.ExecuteNonQuery();
                
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLogErr("StorDocEdit.ExecuteNonQuery: commandText={0}, err={1}", comm.CommandText, e.Message);
            }
            return res;
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

                WebLog.LogClass.WriteToLog("StorDocEdit.AddAutoStorageDoc: delete from Products_StorageDocs id_doc={0}, id_type={1},id_act = {2}, id_branch = {3}, user = {4}, userbranch = {5}", id_doc, id_type, id_act, id_branch, User.Identity.Name, sc.BranchId(User.Identity.Name));

                da.SelectCommand.CommandText = "delete from Products_StorageDocs where id_doc=@id";
                da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id_doc;
                da.SelectCommand.ExecuteNonQuery();

                WebLog.LogClass.WriteToLog("StorDocEdit.AddAutoStorageDoc: delete from Cards_StorageDocs id_doc={0}, id_type={1},id_act = {2}, id_branch = {3}, user = {4}, userbranch = {5}", id_doc, id_type, id_act, id_branch, User.Identity.Name, sc.BranchId(User.Identity.Name));

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

                // отправка в филиал
                if (id_type == 5)
                {
                    //ExecuteQuery(String.Format("select distinct id_BranchCard from V_Cards where (id_stat=2) and ((id_BranchCard_parent={0}) or (id_BranchCard={0}))", id_branch.ToString()), ref ds, trans);
                    // Убираем обособленные офисы 10.12.15 (Рустем)
                    
                    //~!ExecuteQuery(String.Format("select distinct id_BranchCard, Branchs.isolated from V_Cards join Branchs on id_BranchCard=Branchs.id where (id_stat=2) and ((id_BranchCard_parent={0}) or (id_BranchCard={0}))", id_branch.ToString()), ref ds, trans);
                    if (BranchStore.getBranchMainFilial(id_branch, sc.UserAction(User.Identity.Name, Restrictions.Perso))<=0)
                        ExecuteQuery(String.Format("select distinct id_BranchCard, Branchs.isolated from V_Cards join Branchs on id_BranchCard=Branchs.id where (id_stat=2) and ((id_BranchCard_parent={0}) or (id_BranchCard={0}))", id_branch.ToString()), ref ds, trans);
                    else
                        ExecuteQuery(String.Format("select distinct id_BranchCard, Branchs.isolated from V_Cards join Branchs on id_BranchCard=Branchs.id where (id_stat=2) and id_BranchCard={0}", id_branch.ToString()), ref ds, trans);
                    int c_branch = BranchStore.getBranchMainFilial(id_branch, false);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        int id_branch_t = Convert.ToInt32(ds.Tables[0].Rows[i]["id_BranchCard"]);
                        int isolated = Convert.ToInt32(ds.Tables[0].Rows[i]["isolated"]);
                        //20.09.18 Стас закомментарил строчку ниже чтобы заполнить картами документ для родительского филиала
                        //if (isolated != 0 && c_branch == id_branch) continue;
                        InsertCards(id_doc, 2, id_type, id_branch_t, trans);
                    }
                    InsertProductsFromCardsDoc(id_doc, 2, "perso", trans);
                    InsertDopProductsFromCardsDoc(id_doc, 2, "new", trans);
                    //InsertPinsDoc(id_doc, 2, "perso", trans);
                    InsertPinsDocPerso(id_doc, trans);
                }

                if (id_type == 6 || id_type == 1005)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //ExecuteQuery(String.Format("select distinct id_BranchCard from V_Cards where (id_stat=3) and ((id_BranchCard_parent={0}) or (id_BranchCard={0}))", id_branch.ToString()), ref ds, null);
                    //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    //{
                    //    int id_branch_t = Convert.ToInt32(ds.Tables[0].Rows[i]["id_BranchCard"]);
                    //    InsertCards(id_doc, 3, id_type, id_branch_t);
                    //}
                }
                if (id_type == 18)
                {
                    InsertFromDocs_Brak(id_doc, id_act, trans);
                }

                if (id_type == 7 && id_branch == -1)
                {
                    InsertCards(id_doc, 4, id_type, -1, trans);
                }

                if (id_type == 8) //персонализация
                {
                    InsertCards(id_doc, 1, id_type, -1, trans);
                    InsertProductsFromCardsDoc(id_doc, 1, "perso", trans);
                    //InsertPinsDoc(id_doc, 1, "perso", trans);
                    InsertPinsDocPerso(id_doc, trans);
                }

                //~!if (id_type == 9 && isMainFilial())
                //~!{
                    //~!InsertCardsFromDocs(id_doc,id_branch, trans);
                //~!}

                if (id_type == 10)
                {
                    //InsertCards(id_doc, 5, id_type, id_branch);
                    InsertFromDocs(id_doc, id_act, trans);
                }

                if (id_type == 11)
                {
                    //~!if (isMainFilial() == false)
                    //~!{
                        InsertProductsFromStorage(id_doc, -1, "brak", "brak", trans, true);
                        InsertCards(id_doc, 6, id_type, -1, trans);
                        //01.05.2020 продукцию берем из хранилища (2 строки выше)
                        //InsertProductsFromCardsDoc(id_doc, 6, "brak", trans);
                    //InsertPinsDoc1(id_doc, 6, "brak", trans);
                    //~!}
                    //~!else
                    //~!{
                    //~!InsertProductAndCardsFromSendToBank(id_doc,trans);
                    //~!}
                }
                if (id_type == (int)TypeDoc.KillingCard)
                    InsertCards(id_doc, 6, id_type, current_branch, null);
                // отправка на упаковку
                if (id_type == 16)
                {
                    InsertCards(id_doc, 9, id_type, id_branch, trans);
                    InsertProductsFromCardsDoc(id_doc, 9, "perso", trans);
                    InsertDopProductsFromCardsDoc(id_doc, 9, "new", trans);
                    InsertPinsDocPerso(id_doc, trans);
                }
                // прием с упаковки
                if (id_type == 17)
                {
                    InsertFromDocs_Wrap(id_doc, id_act, null);
                    InsertProductsFromCardsDoc(id_doc, 10, "perso", trans);
                    InsertDopProductsFromCardsDoc(id_doc, 10, "new", trans);
                    InsertPinsDocPerso(id_doc, trans);
                }
                // отправка в филиал из филиала
                if (id_type == 19)
                {
                    //int id_deliver = Convert.ToInt32(dListDeliver.SelectedItem.Value);
                    //??????????
                    ds.Clear();
                    
                    //res = ExecuteQuery(String.Format("select id_branch from V_DeliversBranchs_T where (id_branch in (select id_branchCard from Cards where id_stat = 4 and id_branchCurrent={1}) or id_branch in (select id_BranchCard_parent from V_Cards where id_stat=4 and id_branchCurrent={1})) and (id_deliv={0}) order by branch", id_deliver.ToString(), sc.BranchId(User.Identity.Name)), ref ds, null);

//                        ExecuteQuery(String.Format("select distinct id_BranchCard from V_Cards where (id_stat=4) and ((id_BranchCard_parent={0}) or (id_BranchCard={0}))", id_branch.ToString()), ref ds, trans);
                    //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    //{
                        //int id_branch_t = Convert.ToInt32(ds.Tables[0].Rows[i]["id_BranchCard"]);
//                        int id_branch_t = Convert.ToInt32(ds.Tables[0].Rows[i]["id_Branch"]);
                    
                    InsertCards(id_doc, 4, id_type, id_branch, trans);
                    
                    InsertProductsFromCardsDoc(id_doc, 4, "perso", trans);
                    
                    InsertDopProductsFromCardsDoc(id_doc, 4, "new", trans);
                    
                    InsertPinsDocPerso(id_doc, trans);
                    
                }


                if (id_type == (int)TypeDoc.ReceiveToFilialExpertiza)
                {
                    //InsertProductsFromCardsDoc(id_doc, 8, "perso", trans);
                }


                if (id_type == (int)TypeDoc.SendToExpertiza)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 11, "perso", trans);
                }

                if (id_type == (int)TypeDoc.ReceiveToExpertiza)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 12, "perso", trans);
                }
                if (id_type == (int)TypeDoc.Expertiza)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 13, "perso", trans);
                }
                if (id_type == (int)TypeDoc.SendToPodotchet)
                {
                    //InsertProductsFromCardsDoc(id_doc, 4, "perso", trans);
                }
                if (id_type == (int)TypeDoc.ReceiveToPodotchet || id_type == (int)TypeDoc.ToPodotchetFromGoz)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 14, "perso", trans);
                }
                if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
                {
                    //InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 15, "perso", trans);
                }
                if (id_type == (int)TypeDoc.ReturnFromPodotchet || id_type == (int)TypeDoc.FromPodotchetToGoz)
                {
                    #region возврат карт от подотчетного лица
                    using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand comm = conn.CreateCommand())
                        {
                            comm.CommandText = @"insert into Cards_StorageDocs (id_doc, id_card)
                                select @id_doc as id_doc, id from Cards where id_stat=15 and id_person=@id_person";
                            comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                            comm.Parameters.Add("@id_person", SqlDbType.Int).Value = Session["CurrentUserId"];
                            comm.ExecuteNonQuery();
                        }
                        conn.Close();
                    }
                    #endregion
                    //InsertFromDocsExcludeSendClienFromPodotchet(id_doc, id_act, trans);
                }

                if (id_type == (int)TypeDoc.ReceiveFromPodotchet)
                {
                    //InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 17, "perso", trans);
                }
                if (id_type == (int)TypeDoc.WriteOfPodotchet)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 16, "perso", trans);
                }
                if (id_type == (int)TypeDoc.GetBook124)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 18, "perso", trans);
                }
                if (id_type == (int)TypeDoc.ReceiveBook124)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    //InsertProductsFromCardsDoc(id_doc, 20, "perso", trans);
                }
                if (id_type == (int)TypeDoc.GetGoz)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    InsertProductsFromCardsDoc(id_doc, 18, "perso", trans);
                }
                if (id_type == (int)TypeDoc.ReceiveGoz)
                {
                    InsertFromDocs(id_doc, id_act, trans);
                    InsertProductsFromCardsDoc(id_doc, 20, "perso", trans);
                }

            }
            return "";
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
        public string InsertProductAndCardsFromSendToBank(int id_doc, SqlTransaction trans)
        {
                string res = String.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    DataSet ds = new DataSet();
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    string str_type = "";

                    string sqlp = String.Format("select {0} as id_doc,psd.id_prb,0 as cnt_new,0 as cnt_perso,count(psd.id_prb) as cnt_brak" +
                                              " from cards c" +
                                              " inner join cards_storagedocs csd on csd.id_card=c.id" +
                                              " inner join storagedocs sd on csd.id_doc=sd.id and sd.type in (10) and sd.priz_gen>0" +
                                              " inner join products_storagedocs psd on psd.id_doc=csd.id_doc and c.id_prb=psd.id_prb" +
                                              " where c.id_stat=6 and c.id_BranchCurrent={1} group by psd.id_prb",id_doc,branch_main_filial);
                    
                    da.SelectCommand.CommandText = "insert into Products_StorageDocs (id_doc,id_prb,cnt_new,cnt_perso,cnt_brak) " + sqlp;
                    da.SelectCommand.ExecuteNonQuery();
                    ds.Clear();
                    sqlp=String.Format("select {0} as id_doc, c.id as id_card,csd.comment" +
                                       " from cards c" +
                                       " inner join cards_storagedocs csd on csd.id_card=c.id" +
                                       " inner join storagedocs sd on csd.id_doc=sd.id and sd.type in (10) and sd.priz_gen>0" +
                                       " where c.id_stat=6 and c.id_BranchCurrent={1}",id_doc, branch_main_filial);
                    da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card,comment) " + sqlp;
                    da.SelectCommand.ExecuteNonQuery();
                }
                return res;
        }
        public string InsertProductsFromStorage(int id_doc, int id_type, string sost_in, string sost_out, SqlTransaction trans)
        {
            return InsertProductsFromStorage(id_doc,id_type,sost_in,sost_out,trans,false);
        }
        public string InsertProductsFromStorage(int id_doc, int id_type, string sost_in, string sost_out, SqlTransaction trans,bool remove)
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

                    //ExecuteQuery(String.Format("select Storage.id_prb, Storage.cnt_" + sost_in + " as cnt from Storage inner join V_ProductsBanks_T on Storage.id_prb = V_ProductsBanks_T.id_prb where (Storage.cnt_" + sost_in + " > 0) {0}", str_type), ref ds, trans);
                    ExecuteQuery(String.Format("select Storage.id_prb, Storage.cnt_" + sost_in + " as cnt, V_ProductsBanks_T.id_type from Storage inner join V_ProductsBanks_T on Storage.id_prb = V_ProductsBanks_T.id_prb where (Storage.cnt_" + sost_in + " > 0) {0}", str_type), ref ds, trans);
                    
                    // Корректировка для уничтожения
                    //--------------------------------------------------------------
                    /* Пока пришлось убрать для возврата в исходное состояние 19.01.2016
                    if (remove==true)
                    {
                        // Головное подразделение под которым зашли
                        int id_branch = branch_main_filial;
                        if (id_branch < 0) id_branch = 0;
                        DataSet dr = new DataSet();
                        int countpin_minus = 0;
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            int type = Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]);
                            if (type == 1) // карта
                            {
                                int cardscount = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                                int id_prb=Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                                // Выбираем карты для уничтожения отличные от головного
                                dr.Clear();
                                ExecuteQuery("select id,ispin from cards where id_stat=6 and id_prb=" + id_prb.ToString() + " and id_branchCurrent!=" + id_branch.ToString(), ref dr, null);
                                if (dr.Tables.Count > 0 && dr.Tables[0].Rows.Count > 0)
                                {
                                    cardscount -= dr.Tables[0].Rows.Count;
                                    if (cardscount <= 0)
                                    {
                                        ds.Tables[0].Rows.RemoveAt(i);
                                        i--;
                                    }
                                    else
                                    {
                                        ds.Tables[0].Rows[i]["cnt"] = cardscount;
                                    }
                                    for (int j = 0; j < dr.Tables[0].Rows.Count; j++)
                                    {
                                        if (Convert.ToInt32(dr.Tables[0].Rows[j]["ispin"]) > 0) countpin_minus++;
                                    }

                                }
                            }
                        }
                        for (int i = 0; i < ds.Tables[0].Rows.Count && countpin_minus>0; i++)
                        {
                            int type = Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]);
                            if (type == 2) // пин - конверт
                            {
                                int pincount = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                                pincount -= countpin_minus;
                                if (pincount <= 0)
                                {
                                    ds.Tables[0].Rows.RemoveAt(i);
                                    i--;
                                }
                                else
                                {
                                    ds.Tables[0].Rows[i]["cnt"] = pincount;
                                }
                            }
                        }

                    }*/
                    //--------------------------------------------------------------
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        ProductStorageDocsP(id_doc, Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]), Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]), sost_out, trans);
                    }
                }
            return res;
        }
        public string InsertCardsFromDocs(int id_doc,int id_branch, SqlTransaction trans)
        {
            string res = String.Empty;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                DataSet ds = new DataSet();
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card,comment) " +
                                               "select distinct @id_doc,cards.id,Cards_StorageDocs.comment from Cards_StorageDocs " +
                                               "join StorageDocs on StorageDocs.id=Cards_StorageDocs.id_doc " +
                                               "join cards on cards.id=Cards_StorageDocs.id_card " +
                                               "where [type]=10 and id_branch=@id_branch and id_prop=10 and id_stat=4 and priz_gen=1 and " +
                                               "cards.id not in (" +
                                               "select cards.id from Cards_StorageDocs join StorageDocs on StorageDocs.id=Cards_StorageDocs.id_doc join cards on cards.id=Cards_StorageDocs.id_card " +
                                               "where [type]=9 and id_branch=@id_branch and id_prop=10 and (id_stat=5 or id_stat=4))";
                da.SelectCommand.Parameters.Add("@id_branch", SqlDbType.Int).Value = id_branch;
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                da.SelectCommand.ExecuteNonQuery();
            }
            return res;
        }                         
        public string InsertCards(int id_doc, int id_stat, int id_type, int id_branch, SqlTransaction trans)
        {
            string res = String.Empty;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                WebLog.LogClass.WriteToLog("StorDocEdit.InsertCards id_doc={0}, id_type={1},id_stat = {2}, id_branch = {3}, user = {4}, userbranch = {5}", id_doc, id_type, id_stat, id_branch, User.Identity.Name, sc.BranchId(User.Identity.Name));
                DataSet ds = new DataSet();
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;

                string str_branch = "";
                string str_prop = "";
                string str_1 = "";
                if (id_branch != -1)
                {
                    str_branch = String.Format(" and (id_BranchCard={0})", id_branch.ToString());
                    str_1 = String.Format(" and id_branch={0}", id_branch);
                    if (id_type == 19 || id_type == (int)TypeDoc.KillingCard)
                        str_branch += String.Format(" and id_branchCurrent={0}", sc.BranchId(User.Identity.Name));
                }
                if (id_type == 11 && id_branch == -1)
                    str_branch = $" and id_branchCurrent=0";
                if (id_type == 8 || id_type == 5 || id_type == 6 || id_type == 7 || id_type == 19)
                    str_prop = " and id_prop = 1";
                if (id_type == 9 || id_type == 10 || id_type == 11 || id_type == 12)
                {
                    str_prop = " and id_prop <> 1";
                    //Рустем 28.12.15
                    //~!if (isMainFilial()==false && id_type == 11) str_prop += " and id_branchCurrent!=106"; 
                }
                if (id_type == (int)TypeDoc.KillingCard)
                {

                }

                /*
                //ExecuteQuery(String.Format("select id from Cards where id_stat = {0} and (id not in (select id_card from V_CardsTypeDocs where type={1})){2}{3} order by pan", id_stat.ToString(), id_type.ToString(), str_branch, str_prop), ref ds, trans);
                ExecuteQuery(String.Format("select id from Cards where id_stat = {0} and (id not in (select id_card from V_CardsTypeDocs where type={1}{4})){2}{3} order by pan", id_stat.ToString(), id_type.ToString(), str_branch, str_prop, str_1), ref ds, trans);
                ///// *=/
                ExecuteQuery(String.Format("select id from Cards where id_stat = {0}{2}{3} order by pan", id_stat.ToString(), id_type.ToString(), str_branch, str_prop, str_1), ref ds, trans);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    object obj = null;
                    SqlCommand comm = new SqlCommand();
                    comm.Connection = Database.Conn;
                    comm.CommandText = String.Format("select count(*) from V_CardsTypeDocs where id_card={0} and type={1}{2}",Convert.ToInt32(ds.Tables[0].Rows[i]["id"]),id_type, str_1);
                    ExecuteScalar(comm, ref obj, null);
                    if ((int)obj > 0)
                        continue;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card) values (@id_doc,@id_card)";
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.Parameters.Add("@id_card", SqlDbType.Int).Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);
                    da.SelectCommand.ExecuteNonQuery();
                }
                */
                ExecuteQuery(String.Format("select id_card from V_CardsTypeDocs where type={0}{1}",id_type, str_1), ref ds, null);
                Hashtable ht = new Hashtable();
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    if (!ht.ContainsKey(Convert.ToInt32(ds.Tables[0].Rows[i][0])))
                        ht.Add(Convert.ToInt32(ds.Tables[0].Rows[i][0]), "");  
                ds.Clear();
                SqlCommand ins = new SqlCommand();
                ins.Connection = Database.Conn;
                ins.CommandText = "insert into Cards_StorageDocs (id_doc,id_card) values (@id_doc,@id_card)";
                ins.Parameters.Add("@id_doc", SqlDbType.Int);
                ins.Parameters.Add("@id_card", SqlDbType.Int);
                ExecuteQuery(String.Format("select id, id_stat from Cards where id_stat = {0}{2}{3} order by pan", id_stat.ToString(), id_type.ToString(), str_branch, str_prop, str_1), ref ds, trans);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    int id_card = Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);
                    int stat = Convert.ToInt32(ds.Tables[0].Rows[i]["id_stat"]);
                    if (!ht.ContainsKey(id_card) || 
                        (ht.ContainsKey(id_card) && id_type==5 && stat==2) ||
                        (ht.ContainsKey(id_card) && (id_type==5 || id_type==19) && stat==4))
                    {
                        ins.Parameters["@id_doc"].Value = id_doc;
                        ins.Parameters["@id_card"].Value = id_card;
                        ExecuteNonQuery(ins, null);
                    }
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
                //31.10.18 если по карте была экспертиза, то дальнейшее движение она предпринимает без пина
                comm.CommandText = "select id_bank, isPin, (select count(*) from V_Cards_StorageDocs1 where id_card=a.id_card and type=23) as wasExpertiza from V_Cards_StorageDocs a where id_doc=@id_doc";
                comm.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                SqlDataReader dr = comm.ExecuteReader();
                Hashtable ht = new Hashtable();
                while (dr.Read())
                {
                    if (dr["id_bank"] == DBNull.Value)
                        continue;
                    bool wasExpertiza = (Convert.ToInt32(dr["wasExpertiza"]) > 0);
                    if (ht.ContainsKey(Convert.ToInt32(dr["id_bank"])))
                    {
                        if (Convert.ToBoolean(dr["isPin"]) && !wasExpertiza)                                                    
                            ht[Convert.ToInt32(dr["id_bank"])] = (int)ht[Convert.ToInt32(dr["id_bank"])] + 1;
                        
                    }
                    else if (Convert.ToBoolean(dr["isPin"]) && !wasExpertiza)
                        ht.Add(Convert.ToInt32(dr["id_bank"]), (int) 1);
                }

                dr.Close();
                SqlCommand exp = conn.CreateCommand();
                exp.CommandText = "select count(*) from";
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
        public string InsertFromDocsExcludeSendClienFromPodotchet(int id_doc, int id_act, SqlTransaction trans)
        {
            string res = String.Empty;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                DataSet ds = new DataSet();
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card,comment) select @id_doc,id_card,comment from Cards_StorageDocs" +
                                               " where id_doc=@id_act and id_card not in (select id_card from Cards_StorageDocs where id_doc=(select id from StorageDocs where id_act=@id_act and [type]=26))";
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                da.SelectCommand.Parameters.Add("@id_act", SqlDbType.Int).Value = id_act;
                da.SelectCommand.ExecuteNonQuery();
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
                    da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card,comment) select @id_doc,id_card,comment from Cards_StorageDocs where id_doc=@id_act";
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
        public string InsertFromDocs_Brak(int id_doc, int id_act, SqlTransaction trans)
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
                // для отправки обратно в банк при отказе приема ценностей
                // персо превращаются в брак
                da.SelectCommand.CommandText = "insert into Products_StorageDocs (id_doc,id_prb,cnt_brak) select @id_doc,id_prb,cnt_perso from Products_StorageDocs where id_doc=@id_act";
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                da.SelectCommand.Parameters.Add("@id_act", SqlDbType.Int).Value = id_act;
                da.SelectCommand.ExecuteNonQuery();

            }
            return res;
        }
        public string InsertFromDocs_Wrap(int id_doc, int id_act, SqlTransaction trans)
        {
            string res = String.Empty;
            using (SqlDataAdapter da = new SqlDataAdapter())
            {
                DataSet ds = new DataSet();
                da.SelectCommand = new SqlCommand("", conn);
                if (trans != null)
                    da.SelectCommand.Transaction = trans;
                da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                //da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc,id_card) select @id_doc,id_card from Cards_StorageDocs where id_doc=@id_act";
                da.SelectCommand.CommandText = "insert into Cards_StorageDocs (id_doc, id_card) select @id_doc, id_card from V_Cards_StorageDocs where id_doc=@id_act and id_stat=10";
                da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                da.SelectCommand.Parameters.Add("@id_act", SqlDbType.Int).Value = id_act;
                da.SelectCommand.ExecuteNonQuery();

                da.SelectCommand.Parameters.Clear();

                //da.SelectCommand.CommandText = "insert into Products_StorageDocs (id_doc,id_prb,cnt_new,cnt_perso,cnt_brak) select @id_doc,id_prb,cnt_new,cnt_perso,cnt_brak from Products_StorageDocs where id_doc=@id_act";
                //da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                //da.SelectCommand.Parameters.Add("@id_act", SqlDbType.Int).Value = id_act;
                //da.SelectCommand.ExecuteNonQuery();

            }
            return res;
        }

        public string ProductStorageDocsP(int id_doc, int id_prb, int cnt, string sost, SqlTransaction trans)
        {
            string res = String.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter())
                {
                    if (id_prb == 48)
                        id_prb = 48;
                    if (cnt == 74)
                        cnt = 74;
                    da.SelectCommand = new SqlCommand("", conn);
                    if (trans != null)
                        da.SelectCommand.Transaction = trans;
                    da.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    da.SelectCommand.CommandText = "select id from Products_StorageDocs where id_prb=@id_prb and id_doc=@id_doc";
                    da.SelectCommand.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    da.SelectCommand.Parameters.Add("@id_prb", SqlDbType.Int).Value = id_prb;
                    object obj = da.SelectCommand.ExecuteScalar();
                    da.SelectCommand.Parameters.Clear();
                    if (Convert.ToInt32(obj) > 0)
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
                }
            return res;
        }
        protected void WriteToLog(string str)
        {
            lock (Database.lockObjectLog)
            {
                StreamWriter sw = new StreamWriter("c:\\CardPerso\\cardperso.txt", true, System.Text.Encoding.GetEncoding(1251));
                sw.WriteLine(String.Format("{1:dd.MM.yy HH:mm:ss}\t{0}", str, DateTime.Now));
                sw.Close();
            }
        }

    }
}
