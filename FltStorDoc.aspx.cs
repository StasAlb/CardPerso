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
using CardPerso.Administration;

namespace CardPerso
{
    public partial class FltStorDoc : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";

        int branch_main_filial = 0;
        int current_branch = -1;

        ServiceClass sc = new ServiceClass();

        bool Filial = true;
        bool Perso = true;
        bool Accountable = false;
        bool FilialDeliver = true;
      

        protected void Page_Load(object sender, EventArgs e)
        {
            Accountable = SingleQuery.IsAccountable(User.Identity.Name);
            lock (Database.lockObjectDB)
            {
                Filial = sc.UserAction(User.Identity.Name, Restrictions.Filial);
                Perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
                FilialDeliver = sc.UserAction(User.Identity.Name, Restrictions.FilialDeliver);
                current_branch = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(current_branch,Perso);
                
                if (!IsPostBack)
                {
                    ZapCombo();
                    tbProd.Focus();
                }
            }
        }

        private void ZapCombo()
        {
            string sel = "where active=1";
            if (Filial && !Perso)
                sel = sel + " and role_tp=1";
            if (Perso && !Filial)
                sel = sel + " and role_tp=2";
            if (FilialDeliver)
                sel = sel + " or role_tp=3";
            if (branch_main_filial > 0 && branch_main_filial == current_branch)
                sel = sel + " or id=10 or id=11"; // принудительно добавляем возврат из филиала и уничтожение ценностей
            string query = "select id,name from TypeStorageDocs " + sel + " order by id_sort";
            
            ds.Clear();
            //res = Database.ExecuteQuery("select id,name from TypeStorageDocs where active=1 order by id_sort", ref ds, null);
            res = Database.ExecuteQuery(query, ref ds, null);
            dListType.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                int tid = (int)ds.Tables[0].Rows[i]["id"];
                if (Accountable && !(tid == (int) TypeDoc.ReceiveToPodotchet ||
                                     tid == (int) TypeDoc.SendToClientFromPodotchet ||
                                     tid == (int) TypeDoc.ReturnFromPodotchet))
                    continue;
                if (!Accountable && (tid == (int)TypeDoc.ReceiveToPodotchet ||
                                     tid == (int)TypeDoc.SendToClientFromPodotchet ||
                                     tid == (int)TypeDoc.ReturnFromPodotchet))
                    continue;
                dListType.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            }
            dListType.SelectedIndex = 0;
 
            ds.Clear();
            res = Database.ExecuteQuery("select id,Department from Branchs order by department", ref ds, null);
            dListBranch.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListBranch.Items.Add(new ListItem(ds.Tables[0].Rows[i]["Department"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            dListBranch.SelectedIndex = 0;
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ArrayList al = new ArrayList();

                string s = "";

                if (DatePickerStart.DatePickerText != "")
                {
                    try
                    {
                        Convert.ToDateTime(DatePickerStart.DatePickerText);
                    }
                    catch
                    {
                        lbInform.Text = "Неправильно введена дата документа с";
                        DatePickerStart.Focus();
                        return;
                    }
                }
                if (DatePickerEnd.DatePickerText != "")
                {
                    try
                    {
                        Convert.ToDateTime(DatePickerEnd.DatePickerText);
                    }
                    catch
                    {
                        lbInform.Text = "Неправильно введена дата документа по";
                        DatePickerEnd.Focus();
                        return;
                    }
                }

                if (DatePickerStart.DatePickerText != "")
                    al.Add(String.Format("(date_doc>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerStart.SelectedDate));
                if (DatePickerEnd.DatePickerText != "")
                    al.Add(String.Format("(date_doc<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerEnd.SelectedDate));

                string id_type = dListType.SelectedItem.Value;
                if (id_type != "-1")
                    al.Add(String.Format("(type={0})", id_type));

                id_type = dListBranch.SelectedItem.Value;
                if (id_type != "-1")
                    al.Add($"(id_branch={id_type})");

                if (tbProd.Text != "")
                    al.Add(String.Format("(id in (select id_doc from V_Products_StorageDocs where prod_name like [%{0}%]))", tbProd.Text));

                if (tbCard.Text != "")
                {
                    if (tbCard.Text.Trim().Length >= 16)
                        al.Add(
                            $"(id in (select id_doc from V_Cards_StorageDocs where panhash like [{FuncClass.GetHashPan(tbCard.Text.Trim())}]))");
                    else
                        al.Add($"(id in (select id_doc from V_Cards_StorageDocs where pan like [%{tbCard.Text.Trim()}%]))");
                }

                if (tbCreate.Text != "")
                    al.Add(String.Format("(loweredusername like [%{0}%])", tbCreate.Text.Trim()));

                if (chAll.Checked)
                    al.Add(String.Format("(priz_gen=1 or priz_gen=0)", tbCard.Text));
                else
                    al.Add(String.Format("(priz_gen=0)", tbCard.Text));

                if (al.Count > 0)
                {
                    string[] all = Array.CreateInstance(typeof(string), al.Count) as string[];
                    al.CopyTo(all, 0);
                    s = "where " + String.Join(" and ", all);
                }

                Response.Write("<script language=javascript>window.returnValue='" + s + "'; window.close();</script>");
            }
        }
    }
}
