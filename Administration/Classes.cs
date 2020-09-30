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
using OstCard.Data;
using System.Data.SqlClient;
using System.Web.Mvc;

namespace CardPerso.Administration
{
    [Serializable()]
    public class UserClass
    {
        public string FirstName;
        public string SecondName;
        public string LastName;
        public string Position;
        public int BranchId;
        public string Passport;
        public UserClass() 
        {
            FirstName = String.Empty;
            SecondName = String.Empty;
            LastName = String.Empty;
            Position = String.Empty;
            BranchId = -1;
            Passport = String.Empty;
        }
        /// <summary>
        /// должность + фамилия + имя
        /// </summary>
        public string Name1
        {
            get
            {
                return Position + " " + LastName + FirstName;
            }
        }
        public string Fio
        {
            get
            {
                string str = LastName;
                if (FirstName.Trim().Length > 0)
                    str += " " + FirstName.Substring(0,1) + ".";
                if (SecondName.Trim().Length > 0)
                    str += SecondName.Substring(0, 1) + ".";
                return str;
            }
        }
        public string PassportSeries
        {
            get
            {
                string[] p = Passport.Split(' ');
                if (p.Length > 0) return p[0];
                return String.Empty;
            }
        }
        public string PassportNumber
        {
            get
            {
                string[] p = Passport.Split(' ');
                if (p.Length > 1) return p[1];
                return String.Empty;
            }
        }

        public string SetPassport(string series, string number)
        {
            Passport = series.Trim() + " " + number.Trim();
            return Passport;
        }
    }

    public enum Restrictions
    {
        UserRolesEdit = 1,
        Events,
        LibraryView,
        LibraryEdit,
        StorageView,
        CardsView,
        CardsDelete,
        CardsGive,
        PurchaseView,
        PurchaseEdit,
        MovingView,
        MovingEdit,
        Reports,
        AllData,
        ReminderReceiveFilial,
        Filial,
        Perso,
        ConfirmPerso,
        ConfirmSendingValues,
        LibraryOrgEdit,
        ReportGO,
        Transport,
        InformProduction,
        FilialDeliver,
        StorageFilial,
        Report4502,
        AccountBranchView,
        AccountBranchEdit,
        MovingReconfirm,
        ReportKGRT,
        ConfirmKGRT,
        ReceiveToFilialExpertiza,
        SendToExpertiza,
        ReceiveToExpertiza,
        Expertiza,
        SendToPodotchet,
        ReceiveToPodotchet,
        SendToClientFromPodotchet,
        ReturnFromPodotchet,
        ReceiveFromPodotchet,
        WriteOfPodotchet,
        LibraryPodotchet,
        CardKilling, //уничтожение карт в филиале
        ToBook124, // выдать из хранилище по книге 124
        GetBook124, // принять по книге 124
        FromBook124, // вернуть в хранилище по книге 124
        ReceiveBook124, // принять в хранилище по книге 124,
        ReportBook124,
        ToGoz,  // передача в гоз
        GetGoz, // принять в гоз
        FromGoz, // вернуть из гоз
        ReceiveGoz, // принять в хранилище из гоз
        GozToPodotchet,
        PodotchetFromGoz,
        PodotchetToGoz,
        GozFromPodotchet
    }
    public class ServiceClass
    {
        public bool UserAction_v2(string userName, Restrictions action)
        {
            string actions = (string)HttpContext.Current.Session["UserActions"];
            if (actions == null)
            {
                using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand comm = conn.CreateCommand())
                    {
                        comm.CommandText = "select max(id) from Actions";
                        actions = "".PadRight(Convert.ToInt32(comm.ExecuteScalar())+1, '0');
                        comm.CommandText =
                            $"select ActionId from V_UserAction where UserName = '{userName}' order by ActionId";
                        SqlDataReader dr = comm.ExecuteReader();
                        while (dr.Read())
                        {
                            int index = Convert.ToInt32(dr["ActionId"]);
                            actions = actions.Remove(index, 1).Insert(index, "1");
                        }
                        dr.Close();
                        HttpContext.Current.Session["UserActions"] = actions;
                    }
                    conn.Close();
                }
            }
            String currentUserName = (String)HttpContext.Current.Session["CurrentUserName"];
            if (currentUserName != null && userName.ToLower().Equals(currentUserName) == true)
                return ((actions.Length >= (int)action + 1) && actions[(int)action] == '1');
            return false;
        }
        public bool UserAction(string userName, Restrictions action)
        {
            lock (Database.lockObjectDB)
            {
                if (userName.Length == 0)
                    return false;
                if (userName == "AkBarsAdmin" || userName == "UzcardAdmin")
                {
                    if (action == Restrictions.Transport)
                        return false;
                    return true;
                }
                //любая страница начинается здесь, поэтому здесь проверяем на активность коннекта
                if (Database.Conn == null || Database.Conn.State != ConnectionState.Open)
                {
                    Database.Connect(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString);
                    FormsAuthentication.SignOut();
                    FormsAuthentication.RedirectToLoginPage();
                }
                if (Database2.Conn == null || Database2.Conn.State != ConnectionState.Open)
                    Database2.Connect(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString);
                object obj = null;
                string res = Database.ExecuteScalar(String.Format("select count(*) from V_UserAction where UserName='{0}' and ActionId={1}", userName, (int)action), ref obj, null);
                return ((int)obj > 0);
            }
        }
        public string UserGuid(string userName)
        {
            lock (Database.lockObjectDB)
            {
                if (userName == "AkBarsAdmin")
                    return "49BB9584-D482-4E1C-BB80-5A73809E40CD";
                if (userName == "UzcardAdmin")
                    return "C0C75B11-7D2E-408F-865B-BBF472FF8E34";
                object obj = null;
                Database.ExecuteScalar(String.Format("select UserId from aspnet_users where username='{0}'", userName), ref obj, null);
                return (obj == null || obj == DBNull.Value) ? "" : obj.ToString();
            }
        }
        public int UserId(string userName)
        {
            lock (Database.lockObjectDB)
            {
                if (userName == "AkBarsAdmin" || userName == "UzcardAdmin")
                    return 0;
                object obj = null;
                Database.ExecuteScalar(String.Format("select id from aspnet_users where username='{0}'", userName), ref obj, null);
                return (obj == null || obj == DBNull.Value) ? -1 : (int)obj;
            }
        }
        public int BranchId(string userName)
        {
            String branchId=(String)HttpContext.Current.Session["BranchId"];
            if (branchId != null) 
            {
                String currentUserName = (String)HttpContext.Current.Session["CurrentUserName"];
                if (currentUserName != null && userName.ToLower().Equals(currentUserName) == true)
                {
                    return Int32.Parse(branchId);
                }
            }
            lock (Database.lockObjectDB)
            {
                if (userName == "AkBarsAdmin" || userName == "UzcardAdmin")
                    return 0;
                System.Web.Profile.ProfileBase pb = System.Web.Profile.ProfileBase.Create(userName);
                CardPerso.Administration.UserClass uc = (CardPerso.Administration.UserClass)pb.GetPropertyValue("UserData");
                return uc.BranchId;
            }
        }
        
        public string UserFIO(int userId)
        {
            lock (Database.lockObjectDB)
            {
                object obj = null;
                Database.ExecuteScalar($"select UserName from aspnet_users where id={userId}", ref obj, null);
                if (obj == null || obj == DBNull.Value)
                    return "";
                System.Web.Profile.ProfileBase pb = System.Web.Profile.ProfileBase.Create((string)obj);
                CardPerso.Administration.UserClass uc = (CardPerso.Administration.UserClass)pb.GetPropertyValue("UserData");
                return uc.Fio;
            }
        }
        public string UserFIO(string userName)
        {
            lock (Database.lockObjectDB)
            {
                if (userName == "AkBarsAdmin")
                    return "AkBarsAdmin";
                if (userName == "UzcardAdmin")
                    return "UzcardAdmin";
                System.Web.Profile.ProfileBase pb = System.Web.Profile.ProfileBase.Create(userName);
                CardPerso.Administration.UserClass uc = (CardPerso.Administration.UserClass)pb.GetPropertyValue("UserData");
                return uc.Fio;
            }
        }
        public string UserPosition(int userId)
        {
            lock (Database.lockObjectDB)
            {
                object obj = null;
                Database.ExecuteScalar($"select UserName from aspnet_users where id={userId}", ref obj, null);
                if (obj == null || obj == DBNull.Value)
                    return "";
                System.Web.Profile.ProfileBase pb = System.Web.Profile.ProfileBase.Create((string)obj);
                CardPerso.Administration.UserClass uc = (CardPerso.Administration.UserClass)pb.GetPropertyValue("UserData");
                return uc.Position;
            }
        }
        public string UserPosition(string userName)
        {
            lock (Database.lockObjectDB)
            {
                if (userName == "AkBarsAdmin")
                    return "AkBars admin";
                if (userName == "UzcardAdmin")
                    return "Uzcard admin";
                System.Web.Profile.ProfileBase pb = System.Web.Profile.ProfileBase.Create(userName);
                CardPerso.Administration.UserClass uc = (CardPerso.Administration.UserClass)pb.GetPropertyValue("UserData");
                return uc.Position;
            }
        }
        public string UserPassport(string userName)
        {
            lock (Database.lockObjectDB)
            {
                if (userName == "AkBarsAdmin" || userName == "UzcardAdmin")
                    return "0000 000000";
                System.Web.Profile.ProfileBase pb = System.Web.Profile.ProfileBase.Create(userName);
                CardPerso.Administration.UserClass uc = (CardPerso.Administration.UserClass)pb.GetPropertyValue("UserData");
                return uc.Passport;
            }
        }
        public string PassportSeries(string Passport)
        {
            string[] p = Passport.Split(' ');
            if (p.Length == 1 || p.Length == 2)
                return p[0];
            if (p.Length > 2)
                return (p[0] + p[1]).Trim();
            return String.Empty;
            //if (p != null && p.Length > 0) return p[0];
            //return String.Empty;
        }
        public string PassportNumber(string Passport)
        {
            string[] p = Passport.Split(' ');
            if (p.Length == 1)
                return String.Empty;
            if (p.Length == 2)
                return p[1];
            if (p.Length == 3)
                return p[2];
            return String.Empty;
            //if (p != null && p.Length > 1) return p[1];
            //    return String.Empty;
        }

        public void SetUserPassport(string userName, string series, string number)
        {
            SetUserPassport(userName, series.Replace(" ", "") + " " + number.Replace(" ", ""));
        }
        public void SetUserPassport(string userName, string passport)
        {
            lock (Database.lockObjectDB)
            {
                if (userName == "AkBarsAdmin" || userName == "UzcardAdmin")
                    return;
                System.Web.Profile.ProfileBase pb = System.Web.Profile.ProfileBase.Create(userName);
                CardPerso.Administration.UserClass uc = (CardPerso.Administration.UserClass)pb.GetPropertyValue("UserData");
                uc.Passport = passport;
                pb.SetPropertyValue("UserData", uc);
                pb.Save();
            }
        }
    }
}
