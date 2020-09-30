using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Data.SqlClient;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using OstCard.Data;
using CardPerso.Administration;

namespace CardPerso
{
    public partial class OrganizationEdit : System.Web.UI.Page
    {
        private int org_id = -1, p_id = -1, mode = -1;
        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ServiceClass sc = new ServiceClass();
                if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryOrgEdit))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);
                org_id = Convert.ToInt32(Request.QueryString["idO"]);
                p_id = Convert.ToInt32(Request.QueryString["idP"]);
                mode = Convert.ToInt32(Request.QueryString["mode"]);
                Title = (mode == 1) ? "Добавление сотрудника организации" : "Редактирование сотрудника организации";
                if (Page.IsPostBack)
                    return;
                if (mode == 1)
                {
                    DataSet ds = new DataSet();
                    Database.ExecuteQuery("select title,embosstitle from org where id=" + org_id.ToString(), ref ds, null);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        tbTitle.Text = Convert.ToString(ds.Tables[0].Rows[0]["Title"]).Trim();
                        tbEmboss.Text = Convert.ToString(ds.Tables[0].Rows[0]["EmbossTitle"]).Trim();
                    }
                }
                if (mode == 2)
                    LoadOrg();
            }
        }
        private void LoadOrg()
        {
            DataSet ds = new DataSet();
            Database.ExecuteQuery("select * from V_Org where idO="+org_id.ToString() + " and idP=" + p_id.ToString(), ref ds, null);
            if (ds == null || ds.Tables[0].Rows.Count == 0)
                return;
            tbTitle.Text = Convert.ToString(ds.Tables[0].Rows[0]["Title"]).Trim();
            tbEmboss.Text = Convert.ToString(ds.Tables[0].Rows[0]["EmbossTitle"]).Trim();
            tbPerson.Text = Convert.ToString(ds.Tables[0].Rows[0]["Person"]).Trim();
            tbPosition.Text = Convert.ToString(ds.Tables[0].Rows[0]["Position"]).Trim();
            tbPassport.Text = Convert.ToString(ds.Tables[0].Rows[0]["Passport"]).Trim();
            DatePickerPassport.SelectedDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["PDate"]);
            tbPDivision.Text = Convert.ToString(ds.Tables[0].Rows[0]["PDivision"]).Trim();
            tbDoveren.Text = Convert.ToString(ds.Tables[0].Rows[0]["Warrent"]).Trim();
            DatePickerStart.SelectedDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["WStart"]);
            DatePickerEnd.SelectedDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["WEnd"]);
        }
        private bool CheckDate(OstCard.WebControls.DatePicker tb, string lb)
        {
            if (tb.DatePickerText.Trim().Length > 0)
            {
                try
                {
                    Convert.ToDateTime(tb.DatePickerText.Trim());
                }
                catch
                {
                    lInform.Text = "Неправильно введена дата " + lb;
                    tb.Focus();
                    return false;
                }

            }
            else
            {
                lInform.Text = "Неправильно введена дата " + lb;
                tb.Focus();
                return false;
            }
            return true;
        }
        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                /*            if (tbTitle.Text.Trim().Length == 0)
                            {
                                lInform.Text = "Название организации пусто";
                                tbTitle.Focus();
                                return;
                            }
                            if (tbEmboss.Text.Trim().Length == 0)
                            {
                                lInform.Text = "Эмбоссированное название пусто";
                                tbEmboss.Focus();
                                return;
                            }*/
                if (tbPerson.Text.Trim().Length == 0)
                {
                    lInform.Text = "Ответственный пусто";
                    tbPerson.Focus();
                    return;
                }
                if (tbPosition.Text.Trim().Length == 0)
                {
                    lInform.Text = "Должность пусто";
                    tbPosition.Focus();
                    return;
                }
                if (tbPassport.Text.Trim().Length == 0)
                {
                    lInform.Text = "Номер паспорта пусто";
                    tbPassport.Focus();
                    return;
                }
                if (tbPDivision.Text.Trim().Length == 0)
                {
                    lInform.Text = "Паспорт выдан пусто";
                    tbPDivision.Focus();
                    return;
                }
                if (tbDoveren.Text.Trim().Length == 0)
                {
                    lInform.Text = "Номер доверенности пусто";
                    tbDoveren.Focus();
                    return;
                }
                if (!CheckDate(DatePickerPassport, "выдачи паспорта"))
                    return;
                if (!CheckDate(DatePickerStart, "начала действия доверености"))
                    return;
                if (!CheckDate(DatePickerEnd, "окончания действия доверености"))
                    return;

                SqlCommand comm = new SqlCommand();
                if (mode == 1)
                {
                    comm.CommandText = "insert into OrgP (person, position, passport, pdate, pdivision, warrent, wstart, wend, orgid) values (@person, @position, @passport, @pdate, @pdivision, @warrent, @wstart, @wend, @orgid)";
                    comm.Parameters.Add("@orgid", SqlDbType.Int).Value = org_id;
                }
                if (mode == 2)
                {
                    comm.CommandText = "update OrgP set person=@person, position=@position, passport=@passport, pdate=@pdate, pdivision=@pdivision, warrent=@warrent, wstart=@wstart, wend=@wend where id=@id";
                    comm.Parameters.Add("@id", SqlDbType.Int).Value = p_id;
                }
                //            comm.Parameters.Add("@title", SqlDbType.NVarChar, 150).Value = tbTitle.Text.Trim();
                //            comm.Parameters.Add("@embosstitle", SqlDbType.NVarChar, 50).Value = tbEmboss.Text.Trim();
                comm.Parameters.Add("@person", SqlDbType.NVarChar, 150).Value = tbPerson.Text.Trim();
                comm.Parameters.Add("@position", SqlDbType.NVarChar, 30).Value = tbPosition.Text.Trim();
                comm.Parameters.Add("@passport", SqlDbType.NVarChar, 15).Value = tbPassport.Text.Trim();
                comm.Parameters.Add("@pdate", SqlDbType.DateTime).Value = DatePickerPassport.SelectedDate;
                comm.Parameters.Add("@pdivision", SqlDbType.NVarChar, 150).Value = tbPDivision.Text.Trim();
                comm.Parameters.Add("@warrent", SqlDbType.NVarChar, 30).Value = tbDoveren.Text.Trim();
                comm.Parameters.Add("@wstart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                comm.Parameters.Add("@wend", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                Database.ExecuteNonQuery(comm, null);
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
    }
}
