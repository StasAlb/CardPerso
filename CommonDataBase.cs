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
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using CardPerso.Administration;

namespace CardPerso
{
    public class CommonDataBase
    {
        private SqlConnection conn = new SqlConnection();
        
        public CommonDataBase(string nameConnectionString)
        {
            if (nameConnectionString == null) nameConnectionString = "ConString";
            conn.ConnectionString = ConfigurationManager.ConnectionStrings[nameConnectionString].ConnectionString;
            try
            {
                conn.Open();
            }
            catch(Exception e)
            {
                throw new Exception("Произошла ошибка при подключении к базе данных: " + e.Message);
            }
        }

        public DataSet getAccountablePersonsDataset()
        {
            try
            {
                using (SqlDataAdapter dataAdapter = new SqlDataAdapter())
                {
                    DataSet dataSet = new DataSet();
                    dataAdapter.SelectCommand = new SqlCommand("select a.id, secondname + ' ' + firstname + " +
                                                               "case " +
                                                               "when patronymic is null then '' " +
                                                               "else ' ' + patronymic " +
                                                               "end as fio, " +
                                                               "position, " +
                                                               "personnelnumber, UserName " +
                                                               "from AccountablePerson a left join aspnet_Users b on a.UserId=b.id order by secondname, firstname asc", conn);
                    dataAdapter.SelectCommand.CommandTimeout = conn.ConnectionTimeout;
                    dataAdapter.Fill(dataSet);
                    return dataSet;
                }
            }
            catch (Exception e)
            {
                throw new Exception("getAccountablePersonsDataset: " + e.Message);
            }
        }

        public void delAccountablePersonData(int id)
        {
            SqlTransaction tx = null;
            try
            {
                tx = conn.BeginTransaction();
                SqlCommand cmdDelete = new SqlCommand(string.Format("delete from AccountablePersonAccount where id_accountableperson={0}\r\ndelete from AccountablePerson where id={0}", id), conn);
                cmdDelete.Transaction = tx;
                cmdDelete.ExecuteNonQuery();
                tx.Commit();
            }
            catch (Exception e)
            {
                tx.Rollback();
                throw new Exception("delAccountablePersonData: " + e.Message);
            }
        }

        public AccountablePersonData getAccountablePersonData(int id, SqlTransaction trans = null)
        {
            AccountablePersonData apd = new AccountablePersonData();
            if (id != 0)
            {
                SqlCommand cmdSelect = new SqlCommand(string.Format("select a.*, UserName, b.id as UserId from AccountablePerson a left join aspnet_Users b on a.UserId=b.id where a.id = {0}", id), conn);
                cmdSelect.Transaction = trans;
                using (SqlDataReader dr = cmdSelect.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        dr.Read();
                        apd.id = Convert.ToInt32(dr["id"]);
                        apd.secondname = (string)dr["secondname"];
                        apd.firstname = (string)dr["firstname"];
                        if (dr["patronymic"] != DBNull.Value) apd.patronymic = (string)dr["patronymic"];
                        if (dr["birthday"] != DBNull.Value) apd.birthday = Convert.ToDateTime(dr["birthday"]);
                        if (dr["position"] != DBNull.Value) apd.position = (string)dr["position"];
                        if (dr["passport"] != DBNull.Value) apd.passport = (string)dr["passport"];
                        if (dr["dateissue"] != DBNull.Value) apd.dateissue = Convert.ToDateTime(dr["dateissue"]);
                        if (dr["issuedby"] != DBNull.Value) apd.issuedby = (string)dr["issuedby"];
                        if (dr["UserName"] != DBNull.Value) apd.userLogin = (string) dr["UserName"];
                        if (dr["UserId"] != DBNull.Value) apd.userId = (int) dr["UserId"];
                        apd.personnelnumber = (string)dr["personnelnumber"];
                        apd.active = Convert.ToBoolean(dr["active"]);
                    }
                    else throw new Exception("getAccountablePersonData: запись с идентификатором " + id.ToString() + " не найдена");
                }
                cmdSelect = new SqlCommand(string.Format("select * from AccountablePersonAccount where id_accountableperson = {0}", apd.id), conn);
                using (SqlDataReader dr = cmdSelect.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        while (dr.Read()==true)
                        {
                            AccountablePersonAccountData apad = apd.accountdata.FirstOrDefault(r => r.product_type_enum == Convert.ToInt32(dr["product_type_enum"]) &&
                                                                                                  r.account_type_enum == Convert.ToInt32(dr["account_type_enum"]) &&
                                                                                                  r.issafe == Convert.ToBoolean(dr["issafe"]));
                            if (apad != null)
                            {
                                apad.id = Convert.ToInt32(dr["id"]);
                                apad.id_accountableperson = Convert.ToInt32(dr["id_accountableperson"]);
                                apad.account_debet = (string)dr["account_debet"];
                                apad.account_credit = (string)dr["account_credit"];
                            }
                        }
                    }
                }
            }
            return apd;
            
        }

        public int saveAccountablePersonData(AccountablePersonData apd, int branchId)
        {
            SqlTransaction tx = null;
            try
            {
                tx = conn.BeginTransaction();
                SqlCommand cmdAP = new SqlCommand();
                cmdAP.Connection = conn;
                cmdAP.CommandTimeout = conn.ConnectionTimeout;
                cmdAP.Transaction = tx;
                SqlCommand cmdAPA = new SqlCommand();
                cmdAPA.Connection = conn;
                cmdAPA.CommandTimeout = conn.ConnectionTimeout;
                cmdAPA.Transaction = tx;
                //@secondname,@firstname,@patronymic,@birthday,@position,@passport,@dateissue,@issuedby,@personnelnumber,@active 
                //secondname,firstname,patronymic,birthday,position,passport,dateissue,issuedby,personnelnumber,active 
                if (apd.id == 0)
                {
                    cmdAP.CommandText = "insert into AccountablePerson " +
                                              "(secondname,firstname,patronymic,birthday,position,passport,dateissue,issuedby,personnelnumber,active) " +
                                              "values(@secondname,@firstname,@patronymic,@birthday,@position,@passport,@dateissue,@issuedby,@personnelnumber,@active) select @@identity as newid";
                }
                else
                {
                    cmdAP.CommandText = "update AccountablePerson " +
                                        "set " +
                                        "secondname=@secondname," +
                                        "firstname=@firstname," +
                                        "patronymic=@patronymic," +
                                        "birthday=@birthday," +
                                        "position=@position," +
                                        "passport=@passport," +
                                        "dateissue=@dateissue," +
                                        "issuedby=@issuedby," +
                                        "personnelnumber=@personnelnumber," +
                                        "active=@active " +
                                        "where id=@id";
                    cmdAP.Parameters.Add("id", SqlDbType.Int).Value = apd.id;
                }
                cmdAP.Parameters.Add("secondname", SqlDbType.VarChar, 50).Value = apd.secondname.Trim();
                cmdAP.Parameters.Add("firstname", SqlDbType.VarChar, 20).Value = apd.firstname.Trim();
                cmdAP.Parameters.Add("patronymic", SqlDbType.VarChar, 30).Value = (apd.patronymic.Trim().Length > 0) ? apd.patronymic.Trim() : Convert.DBNull;
                cmdAP.Parameters.Add("birthday", SqlDbType.DateTime).Value = (apd.birthday != DateTime.MinValue) ? apd.birthday : Convert.DBNull;
                cmdAP.Parameters.Add("position", SqlDbType.VarChar, 40).Value = (apd.position.Trim().Length > 0) ? apd.position : Convert.DBNull;
                cmdAP.Parameters.Add("passport", SqlDbType.VarChar, 20).Value = (apd.passport.Trim().Length > 0) ? apd.passport.Trim() : Convert.DBNull;
                cmdAP.Parameters.Add("dateissue", SqlDbType.DateTime).Value = (apd.dateissue != DateTime.MinValue) ? apd.dateissue : Convert.DBNull;
                cmdAP.Parameters.Add("issuedby", SqlDbType.VarChar, 100).Value = (apd.issuedby.Trim().Length > 0) ? apd.issuedby.Trim() : Convert.DBNull;
                cmdAP.Parameters.Add("personnelnumber", SqlDbType.VarChar, 20).Value = apd.personnelnumber.Trim();
                cmdAP.Parameters.Add("active", SqlDbType.Bit).Value = apd.active;
                if (apd.id == 0)
                {
                    object obj = cmdAP.ExecuteScalar();
                    apd.id = Convert.ToInt32(obj);
                    #region добавляем подотчетное лицо как обычного пользователя кардперсо
                    SqlCommand checkLogin = conn.CreateCommand();
                    checkLogin.Transaction = tx;
                    checkLogin.CommandText = "select count(*) from aspnet_Users where username=@username";
                    checkLogin.Parameters.Add("@username", SqlDbType.NVarChar, 255).Value = apd.userLogin;
                    obj = checkLogin.ExecuteScalar();
                    if ((int) obj > 0)
                        throw new Exception("Данный логин занят");

                    Guid guid = Guid.NewGuid();
                    SqlCommand user = conn.CreateCommand();
                    user.Transaction = tx;
                    user.CommandText = "insert into aspnet_Users (applicationid, userid, username, loweredusername, isanonymous, lastactivitydate, activepassword) values ('FDD44FA6-F1C8-42AA-9082-E5B22438E20F', @userid, @username, @lower, 0, @lastactivity, 0) select @@identity";
                    user.Parameters.Add("@userid", SqlDbType.UniqueIdentifier).Value = guid;
                    user.Parameters.Add("@username", SqlDbType.NVarChar, 255).Value = apd.userLogin;
                    user.Parameters.Add("@lower", SqlDbType.NVarChar, 255).Value = apd.userLogin.ToLower();
                    user.Parameters.Add("@lastactivity", SqlDbType.DateTime).Value = DateTime.Now;

                    SqlCommand inRoles = conn.CreateCommand();
                    inRoles.Transaction = tx;
                    inRoles.CommandText = $"insert into aspnet_UsersInRoles (userid, roleid) values (@userid, '{(string)ConfigurationManager.AppSettings["AccountablePersons"]}')";
                    inRoles.Parameters.Add("@userid", SqlDbType.UniqueIdentifier).Value = guid;
                    SqlCommand profile = conn.CreateCommand();
                    profile.Transaction = tx;
                    profile.CommandText = "insert into aspnet_Profile (userid, propertynames, propertyvaluesstring,propertyvaluesbinary,lastupdateddate) values (@userid, @property,@pvalue,@bvalue,@last)";
                    profile.Parameters.Add("@userid", SqlDbType.UniqueIdentifier).Value = guid;
//                    string data = Encoding.ASCII.GetString(Utilities.AHex2Bin("3C3F786D6C2076657273696F6E3D22312E302220656E636F64696E673D227574662D3136223F3E0D0A3C55736572436C61737320786D6C6E733A7873693D22687474703A2F2F7777772E77332E6F72672F323030312F584D4C536368656D612D696E7374616E63652220786D6C6E733A7873643D22687474703A2F2F7777772E77332E6F72672F323030312F584D4C536368656D61223E0D0A20203C46697273744E616D653E232346495253544E414D4523233C2F46697273744E616D653E0D0A20203C5365636F6E644E616D653E23235345434F4E444E414D4523233C2F5365636F6E644E616D653E0D0A20203C4C6173744E616D653E23234C4153544E414D4523233C2F4C6173744E616D653E0D0A20203C506F736974696F6E3E2323504F5323233C2F506F736974696F6E3E0D0A20203C4272616E636849643E23234252414E434823233C2F4272616E636849643E0D0A3C2F55736572436C6173733E0D0A"));
                    string data = Encoding.ASCII.GetString(Utilities.AHex2Bin("3C3F786D6C2076657273696F6E3D22312E302220656E636F64696E673D227574662D3136223F3E0D0A3C55736572436C61737320786D6C6E733A7873693D22687474703A2F2F7777772E77332E6F72672F323030312F584D4C536368656D612D696E7374616E63652220786D6C6E733A7873643D22687474703A2F2F7777772E77332E6F72672F323030312F584D4C536368656D61223E0D0A20203C46697273744E616D653E232346495253544E414D4523233C2F46697273744E616D653E0D0A20203C5365636F6E644E616D653E23235345434F4E444E414D4523233C2F5365636F6E644E616D653E0D0A20203C4C6173744E616D653E23234C4153544E414D4523233C2F4C6173744E616D653E0D0A20203C506F736974696F6E3E2323504F5323233C2F506F736974696F6E3E0D0A20203C4272616E636849643E23234252414E434823233C2F4272616E636849643E0D0A20203C50617373706F72743E232350415353504F525423233C2F50617373706F72743E0D0A3C2F55736572436C6173733E0D0A"));
                    //@"<?xml version = ""1.0"" encoding = ""utf-16""?>< UserClass xmlns:xsi = ""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd = ""http://www.w3.org/2001/XMLSchema"" >< FirstName > ##FIRSTNAME## </ FirstName >< SecondName > ##SECONDNAME## </ SecondName >< LastName > ##LASTNAME## </ LastName >< Position > ##POS## </ Position >< BranchId > ##BRANCH## </ BranchId >< Passport > ##PASSPORT## </ PASSPORT ></ UserClass >";
                    data = data.Replace("##FIRSTNAME##", apd.firstname);
                    data = data.Replace("##SECONDNAME##", apd.patronymic);
                    data = data.Replace("##LASTNAME##", apd.secondname);
                    data = data.Replace("##POS##", apd.position);
                    data = data.Replace("##BRANCH##", branchId.ToString());
                    data = data.Replace("##PASSPORT##", apd.PassportSeries.Replace(" ", "") + " " + apd.PassportNumber.Replace(" ", ""));

                    Random r = new Random(DateTime.Now.Millisecond);
                    byte[] byteSalt = new byte[16];
                    r.NextBytes(byteSalt);
                    byte[] bytePass = Encoding.Unicode.GetBytes("123");
                    byte[] byteResult = new byte[byteSalt.Length + bytePass.Length];
                    System.Buffer.BlockCopy(byteSalt, 0, byteResult, 0, byteSalt.Length);
                    System.Buffer.BlockCopy(bytePass, 0, byteResult, byteSalt.Length, bytePass.Length);

                    HashAlgorithm ha = HashAlgorithm.Create("SHA1");

                    profile.Parameters.Add("@property", SqlDbType.NText).Value = "UserData:S:0:" + data.Length.ToString() + ":"; ;
                    profile.Parameters.Add("@pvalue", SqlDbType.NText).Value = data;
                    profile.Parameters.Add("@bvalue", SqlDbType.Image).Value = new byte[] { };
                    profile.Parameters.Add("@last", SqlDbType.DateTime).Value = DateTime.Now;

                    SqlCommand member = conn.CreateCommand();
                    member.Transaction = tx;
                    member.CommandText = "insert into aspnet_membership (applicationid, userid, password, passwordformat, passwordsalt,email, loweredemail, isapproved, islockedout, createdate, lastlogindate,lastpasswordchangeddate,lastlockoutdate, FailedPasswordAttemptCount,[FailedPasswordAttemptWindowStart],[FailedPasswordAnswerAttemptCount],[FailedPasswordAnswerAttemptWindowStart]) values ('FDD44FA6-F1C8-42AA-9082-E5B22438E20F',@userid,@password,1,@passwordsalt,@email,@lowere,1,0,@cdate,@cdate,@cdate,@cdate,0,@cdate,0,@cdate)";
                    member.Parameters.Add("@userid", SqlDbType.UniqueIdentifier).Value = guid;
                    member.Parameters.Add("@password", SqlDbType.NVarChar, 128).Value = Convert.ToBase64String(ha.ComputeHash(byteResult));
                    member.Parameters.Add("@passwordsalt", SqlDbType.NVarChar, 128).Value = Convert.ToBase64String(byteSalt);
                    member.Parameters.Add("@email", SqlDbType.NVarChar, 256).Value = DBNull.Value;
                    member.Parameters.Add("@lowere", SqlDbType.NVarChar, 256).Value = DBNull.Value;
                    member.Parameters.Add("@cdate", SqlDbType.DateTime).Value = DateTime.Now;

                    object uid = user.ExecuteScalar();
                    inRoles.ExecuteNonQuery();
                    profile.ExecuteNonQuery();
                    member.ExecuteNonQuery();

                    SqlCommand comm = conn.CreateCommand();
                    comm.Transaction = tx;
                    comm.CommandText = $"update AccountablePerson set UserId={uid} where id={apd.id}";
                    comm.ExecuteNonQuery();
                    #endregion
                }
                else
                {
                    #region редактируем подотчетное лицо как обычного пользователя кардперсо

                    using (SqlConnection conn1 = new SqlConnection())
                    {
                        conn1.ConnectionString = FuncClass.ConnectionString;
                        conn1.Open();
                        SqlCommand checkLogin = conn1.CreateCommand();
                        checkLogin.CommandText =
                            "select count(*) from aspnet_Users where username=@username and id<>(select userid from accountableperson where id=@id)";
                        checkLogin.Parameters.Add("@username", SqlDbType.NVarChar, 255).Value = apd.userLogin;
                        checkLogin.Parameters.Add("@id", SqlDbType.Int).Value = apd.id;
                        object obj = checkLogin.ExecuteScalar();
                        if ((int) obj > 0)
                            throw new Exception("Данный логин занят");
                        SqlCommand ccc = conn1.CreateCommand();
                        ccc.CommandText = "select userid from aspnet_users where id=(select userid from accountableperson where id=@uid)";
                        ccc.Parameters.Add("@uid", SqlDbType.Int).Value = apd.id;
                        object id = ccc.ExecuteScalar();
                        if (id != null && id != DBNull.Value)
                        {

                            SqlCommand user = conn1.CreateCommand();
                            user.CommandText =
                                "update aspnet_Users set username=@username, loweredusername=@lower, lastactivitydate=@lastactivity where userid=@uid";
                            user.Parameters.Add("@uid", SqlDbType.UniqueIdentifier).Value = (Guid)id;
                            user.Parameters.Add("@username", SqlDbType.NVarChar, 255).Value = apd.userLogin;
                            user.Parameters.Add("@lower", SqlDbType.NVarChar, 255).Value = apd.userLogin.ToLower();
                            user.Parameters.Add("@lastactivity", SqlDbType.DateTime).Value = DateTime.Now;

                            SqlCommand profile = conn1.CreateCommand();
                            profile.CommandText =
                                "update aspnet_Profile set propertynames=@property, propertyvaluesstring=@pvalue where userid=@uid";
                            string data = Encoding.ASCII.GetString(Utilities.AHex2Bin(
                                "3C3F786D6C2076657273696F6E3D22312E302220656E636F64696E673D227574662D3136223F3E0D0A3C55736572436C61737320786D6C6E733A7873693D22687474703A2F2F7777772E77332E6F72672F323030312F584D4C536368656D612D696E7374616E63652220786D6C6E733A7873643D22687474703A2F2F7777772E77332E6F72672F323030312F584D4C536368656D61223E0D0A20203C46697273744E616D653E232346495253544E414D4523233C2F46697273744E616D653E0D0A20203C5365636F6E644E616D653E23235345434F4E444E414D4523233C2F5365636F6E644E616D653E0D0A20203C4C6173744E616D653E23234C4153544E414D4523233C2F4C6173744E616D653E0D0A20203C506F736974696F6E3E2323504F5323233C2F506F736974696F6E3E0D0A20203C4272616E636849643E23234252414E434823233C2F4272616E636849643E0D0A20203C50617373706F72743E232350415353504F525423233C2F50617373706F72743E0D0A3C2F55736572436C6173733E0D0A"));
                            //@"<?xml version = ""1.0"" encoding = ""utf-16""?>< UserClass xmlns:xsi = ""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd = ""http://www.w3.org/2001/XMLSchema"" >< FirstName > ##FIRSTNAME## </ FirstName >< SecondName > ##SECONDNAME## </ SecondName >< LastName > ##LASTNAME## </ LastName >< Position > ##POS## </ Position >< BranchId > ##BRANCH## </ BranchId >< Passport > ##PASSPORT## </ PASSPORT ></ UserClass >";
                            data = data.Replace("##FIRSTNAME##", apd.firstname);
                            data = data.Replace("##SECONDNAME##", apd.patronymic);
                            data = data.Replace("##LASTNAME##", apd.secondname);
                            data = data.Replace("##POS##", apd.position);
                            data = data.Replace("##BRANCH##", branchId.ToString());
                            data = data.Replace("##PASSPORT",
                                apd.PassportSeries.Replace(" ", "") + " " + apd.PassportNumber.Replace(" ", ""));

                            profile.Parameters.Add("@uid", SqlDbType.UniqueIdentifier).Value = (Guid)id;
                            profile.Parameters.Add("@property", SqlDbType.NText).Value =
                                "UserData:S:0:" + data.Length.ToString() + ":";
                            profile.Parameters.Add("@pvalue", SqlDbType.NText).Value = data;
                            //profile.Parameters.Add("@last", SqlDbType.DateTime).Value = DateTime.Now;
                            user.ExecuteNonQuery();
                            profile.ExecuteNonQuery();
                        }

                        conn1.Close();
                    }
                    #endregion
                    cmdAP.ExecuteNonQuery();
                }
                cmdAPA.Parameters.Add("id", SqlDbType.Int);
                cmdAPA.Parameters.Add("id_accountableperson", SqlDbType.Int);
                cmdAPA.Parameters.Add("product_type_enum", SqlDbType.Int);
                cmdAPA.Parameters.Add("account_type_enum", SqlDbType.Int);
                cmdAPA.Parameters.Add("account_debet", SqlDbType.VarChar, 20);
                cmdAPA.Parameters.Add("account_credit", SqlDbType.VarChar, 20);
                cmdAPA.Parameters.Add("issafe", SqlDbType.Bit);

                cmdAPA.Parameters["id_accountableperson"].Value = apd.id;

                for (int i = 0; i < apd.accountdata.Length; i++)
                {
                    cmdAPA.Parameters["id"].Value = apd.accountdata[i].id;
                    cmdAPA.Parameters["product_type_enum"].Value = apd.accountdata[i].product_type_enum;
                    cmdAPA.Parameters["account_type_enum"].Value = apd.accountdata[i].account_type_enum;
                    cmdAPA.Parameters["account_debet"].Value = (string)apd.accountdata[i].account_debet.Trim();
                    cmdAPA.Parameters["account_credit"].Value = (string)apd.accountdata[i].account_credit.Trim();
                    cmdAPA.Parameters["issafe"].Value = apd.accountdata[i].issafe;
                    if (apd.accountdata[i].id > 0)
                    {
                        
                        cmdAPA.CommandText = "update AccountablePersonAccount set " +
                                             "account_debet=@account_debet," +
                                             "account_credit=@account_credit " +
                                             "where id=@id and id_accountableperson=@id_accountableperson and product_type_enum=@product_type_enum and " +
                                             "account_type_enum=@account_type_enum and issafe=@issafe select @id as oldid";
                    }
                    else
                    {
                        cmdAPA.CommandText = "insert into AccountablePersonAccount " +
                                             "(id_accountableperson,product_type_enum,account_type_enum,account_debet,account_credit,issafe) " +
                                             "values(@id_accountableperson,@product_type_enum,@account_type_enum,@account_debet,@account_credit,@issafe) select @@identity as newid";
                    }
                    object obj = cmdAPA.ExecuteScalar();
                    if (apd.accountdata[i].id < 1)
                    {
                        apd.accountdata[i].id = Convert.ToInt32(obj);
                    }

                }
                tx.Commit();
                return apd.id;
            }
            catch (Exception e)
            {
                tx.Rollback();
                throw new Exception("saveAccountablePersonsData: " + e.Message);
            }
        }
    }
}
