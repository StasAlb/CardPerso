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
using System.Data.SqlClient;
using OstCard.Data;
using CardPerso.Administration;


namespace CardPerso
{
    public partial class ConfField : System.Web.UI.Page
    {
        string type_tbl = "";
        ServiceClass sc = new ServiceClass();

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                type_tbl = Request.QueryString["type"];
                if (!IsPostBack) ZapFields();
            }
        }

        private void ZapFields()
        {
             if (type_tbl == "storage")
             {
                
                chFields.Items.Add(new ListItem("Банк", "bank_name"));
                chFields.Items.Add(new ListItem("Бин", "bin"));
                chFields.Items.Add(new ListItem("Минимум", "min_cnt"));
                chFields.Items.Add(new ListItem("Новые", "cnt_new"));
                //chFields.Items.Add(new ListItem("На руках", "cnt_wrk"));
                chFields.Items.Add(new ListItem("Персо", "cnt_perso"));
                chFields.Items.Add(new ListItem("Брак", "cnt_brak"));
               // chFields.Items.Add(new ListItem("Уничтожение", "cnt_del"));
            }

            if (type_tbl == "card")
            {                
                chFields.Items.Add(new ListItem("Банк", "bank_name"));                
                chFields.Items.Add(new ListItem("Дата начала", "dateStart"));
                chFields.Items.Add(new ListItem("Дата окончания", "dateEnd"));
                chFields.Items.Add(new ListItem("Паспорт", "passport"));
                chFields.Items.Add(new ListItem("Номер счета", "account"));
                chFields.Items.Add(new ListItem("Дата изготовления", "dateProd"));                
                chFields.Items.Add(new ListItem("Курьерская служба", "courier"));
                chFields.Items.Add(new ListItem("Номер накладной", "invoice"));
                chFields.Items.Add(new ListItem("Дата передачи", "date_courier"));
                chFields.Items.Add(new ListItem("Филиал выпуска", "DepBranchInit"));
                chFields.Items.Add(new ListItem("Филиал отправки", "DepBranchCard"));
                chFields.Items.Add(new ListItem("Филиал отправки (полный)", "BranchCardTransport"));
                chFields.Items.Add(new ListItem("Дата получения", "dateReceipt"));
                chFields.Items.Add(new ListItem("Дата выдачи", "dateClient"));
                chFields.Items.Add(new ListItem("Выдавший сотрудник", "clientWorker"));
                chFields.Items.Add(new ListItem("Дата об уничтожение", "dateSendTerminate"));
                chFields.Items.Add(new ListItem("Причина отметки", "Comment"));
                chFields.Items.Add(new ListItem("Дата получения (филиал)", "dateGetTerminate"));
                chFields.Items.Add(new ListItem("Дата уничтожения", "dateTerminated"));
            }

            if (type_tbl == "purchase_dog")
            {
                chFields.Items.Add(new ListItem("Номер", "number_dog"));
                chFields.Items.Add(new ListItem("Дата", "date_dog"));
                chFields.Items.Add(new ListItem("Дата хранилище", "date_stor"));
                chFields.Items.Add(new ListItem("Поставщик", "supplier"));
                chFields.Items.Add(new ListItem("Изготовитель", "manuf"));
                chFields.Items.Add(new ListItem("Дата выписки", "date_record"));
                chFields.Items.Add(new ListItem("Отметка выписки", "comment"));
            }

            if (type_tbl == "purchase_product")
            {
                chFields.Items.Add(new ListItem("Наименование", "prod_name"));
                chFields.Items.Add(new ListItem("Банк", "bank_name"));
                chFields.Items.Add(new ListItem("Кол-во", "cnt"));
                chFields.Items.Add(new ListItem("Цена", "price"));
                chFields.Items.Add(new ListItem("Сумма", "summa"));
            }

            if (type_tbl == "stordoc")
            {
                chFields.Items.Add(new ListItem("Номер", "number_doc"));
                chFields.Items.Add(new ListItem("Дата", "date_doc"));
                chFields.Items.Add(new ListItem("Тип документа", "type_name"));
                chFields.Items.Add(new ListItem("Филиал", "branch"));
                chFields.Items.Add(new ListItem("Состояние", "gen"));
                chFields.Items.Add(new ListItem("Комментарий", "comment"));
            }

            for (int i = 0; i < chFields.Items.Count; i++)
                chFields.Items[i].Selected = true;

            SelectFields();
        }

        private void SelectFields()
        {
            string s_fld = Database.GetFiledsByUser(sc.UserId(User.Identity.Name), type_tbl, null);
            if (s_fld!="")
            {
                string[] ar_fld = s_fld.Split(Convert.ToChar(","));
                for (int k = 0; k < ar_fld.Count(); k++)
                    chFields.Items.FindByValue(ar_fld[k]).Selected = false;
            }

        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ArrayList ar_fld = new ArrayList();
                string fld = "";

                for (int i = 0; i < chFields.Items.Count; i++)
                    if (!chFields.Items[i].Selected) ar_fld.Add(chFields.Items[i].Value);

                if (ar_fld.Count == chFields.Items.Count) return;

                if (ar_fld.Count > 0)
                {
                    string[] s_fld = Array.CreateInstance(typeof(string), ar_fld.Count) as string[];
                    ar_fld.CopyTo(s_fld, 0);
                    fld = String.Join(",", s_fld);
                }
                else fld = "";

                SqlCommand sqCom = new SqlCommand();

                int id_uf = Database.GetUserFieldsId(sc.UserId(User.Identity.Name), type_tbl, null);

                if (id_uf == 0)
                {
                    sqCom.CommandText = "insert into UserFields (UserId,tbl_name,fld_name) values(@UserId,@tbl_name,@fld_name)";
                    sqCom.Parameters.Add("@UserId", SqlDbType.Int).Value = sc.UserId(User.Identity.Name);
                    sqCom.Parameters.Add("@tbl_name", SqlDbType.VarChar, 30).Value = type_tbl;
                    sqCom.Parameters.Add("@fld_name", SqlDbType.VarChar, 300).Value = fld;
                }
                else
                {
                    sqCom.CommandText = "update UserFields set fld_name=@fld_name where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_uf;
                    sqCom.Parameters.Add("@fld_name", SqlDbType.VarChar, 300).Value = fld;
                }
                Database.ExecuteNonQuery(sqCom, null);
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
    }
}
