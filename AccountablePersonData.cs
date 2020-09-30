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


namespace CardPerso
{
    public class AccountablePersonData
    {
        public const int COUNTOPERATION = 5;
        public const string ACCOUNTSTORAGE = "91202";
        public const string ACCOUNTSAFE = "91203";
        public const string ACCOUNTCLIENT = "99999";
        public int id { get; set; }
        public string secondname { get; set; }
        public string firstname  { get; set; } 
        public string patronymic { get; set; }
        public DateTime birthday { get; set; }
        public string position { get; set; }
        public string passport { get; set; }
        public DateTime dateissue { get; set; }
        public string issuedby { get; set; }
        public string personnelnumber { get; set; }
        public bool active { get; set; }
        public string userLogin { get; set; }
        /// <summary>
        /// id пользователя aspnet user
        /// </summary>
        public int userId { get; set; }
        public AccountablePersonAccountData[]  accountdata { get; set; }

        public string PassportSeries
        {
            get
            {
                string[] p = passport.Split(' ');
                if (p != null && p.Length > 0) return p[0];
                return String.Empty;
            }
        }
        public string PassportNumber
        {
            get
            {
                string[] p = passport.Split(' ');
                if (p != null && p.Length > 1) return p[1];
                return String.Empty;
            }
        }

      public AccountablePersonData()
      {
          accountdata = new AccountablePersonAccountData[(Enum.GetNames(typeof(BaseProductType)).Length - 1) * COUNTOPERATION];
                   
          secondname = "";
          firstname = "";
          patronymic = "";
          birthday = DateTime.MinValue;
          position = "";
          passport = "";
          dateissue = DateTime.MinValue;
          issuedby = "";
          personnelnumber = "";
          active = true;
          userLogin = "";
          userId = 0;
        
          foreach (var value in Enum.GetValues(typeof(BaseProductType))) 
          {
              if ((int)value == 0) continue;
              int index = ((int)value) - 1;
              for (int i = 0; i < COUNTOPERATION; i++)
              {
                  accountdata[index * COUNTOPERATION + i] = new AccountablePersonAccountData();
                  AccountablePersonAccountData ad = accountdata[index * COUNTOPERATION + i];
                  ad.product_type_enum = (int)value;
                  if (i < 2)
                  {
                      if (i == 1)
                      {
                          ad.account_credit = ACCOUNTSAFE;
                          ad.issafe = true;
                      }
                      else
                      {
                          ad.account_credit = ACCOUNTSTORAGE;
                          ad.issafe = false;
                      }
                      ad.account_debet = "91203???????????????";
                      ad.account_type_enum = (int)AccountBranchType.In;
                    
                  }
                  if (i == 2)
                  {
                      ad.account_credit = "91203???????????????";
                      ad.account_debet = ACCOUNTCLIENT;
                      ad.issafe = false;
                      ad.account_type_enum = (int)AccountBranchType.Out;
                  }
                  if (i > 2 && i < 5)
                  {
                      if (i == 4)
                      {
                          ad.account_debet = ACCOUNTSAFE;
                          ad.issafe = true;
                      }
                      else
                      {
                          ad.account_debet = ACCOUNTSTORAGE;
                          ad.issafe = false;
                      }
                      ad.account_credit = "91203???????????????";
                      ad.account_type_enum = (int)AccountBranchType.Return;
                  
                  }
              }
          }
      }
    }

    public class AccountablePersonAccountData
    {
        public int id { get; set; }
        public int id_accountableperson { get; set; }
        public int product_type_enum { get; set; }
        public int account_type_enum { get; set; }
        public string account_debet { get; set; }
        public string account_credit { get; set; }
        public bool issafe { get; set; }
    }
}
