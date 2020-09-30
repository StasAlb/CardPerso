using System;
using System.Data;
using System.Web.UI.WebControls;
using OstCard.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using CardPerso.Administration;
using WebLog;


namespace CardPerso
{
    public enum TypeDoc
    {
        CardFromStorage = 1,
        PinFromStorage = 2,
        CardToStorage = 3,
        PinToStorage = 4,
        SendToFilial = 5,
        ReceiveToFilial = 6,
        SendToClient = 7,
        PersoCard = 8,
        SendToBank = 9,
        ReceiveToBank = 10,
        DeleteBrak = 11,
        ReturnToFilial=12,
        FilialFilial = 13,
        Expendables = 14,
        Reklama = 15,
        ToWrapping = 16,
        FromWrapping = 17,
        DontReceiveToFilial = 18,
        SendToFilialFilial = 19,
        ReceiveToFilialExpertiza = 20,
        SendToExpertiza = 21,
        ReceiveToExpertiza = 22,
        Expertiza = 23,
        SendToPodotchet = 24,
        ReceiveToPodotchet = 25,
        SendToClientFromPodotchet = 26,
        ReturnFromPodotchet = 27,
        ReceiveFromPodotchet = 28,
        WriteOfPodotchet = 29,
        KillingCard = 30,
        // ограничить их 30, иначе пойдут пересечения с TypeFormDoc (вот так там по идиотски сделано :(((
        #warning посмотреть что с этим можно сделать
        ToBook124 = 1000,
        GetBook124 = 1001,
        FromBook124 = 1002,
        ReceiveBook124 = 1003,
        SendToClientService = 1004,
        ReceiveToFilialPacket = 1005,
        ToGoz = 1006,
        GetGoz = 1007,
        FromGoz = 1008,
        ReceiveGoz = 1009,
        FromGozToPodotchet,
        ToPodotchetFromGoz,
        FromPodotchetToGoz,
        ToGozFromPodotchet
    }
    public enum CardStatus : int
    {
        Undefined = 0,
        Proccess = 1,
        Storage = 2,
        CourierFilial = 3,
        Filial = 4,
        CourierBank = 5,
        Killing = 6,
        Killed = 7,
        Client = 8,
        StorageWrap = 9,
        Wrapping = 10,
        FilialExpertiza = 11,
        ExpertizaInWay = 12,
        StorageExpertiza = 13,
        FilialAccountable = 14,
        Accountable = 15,
        AccountableClient = 16,
        AccountableFilial = 17,
        FilialBook124 = 18,
        Book124 = 19,
        Book124Filial = 20,
        ExpectToClient = 21,
        FilialPacket = 22,
        FilialGoz = 23,
        Goz = 24,
        GozFilial = 25,
        GozAccountable = 26,
        AccountableGoz = 27
    }
    public enum TypeFormDoc
    {
        OfficeNote = 31,
        Register = 32,
        SortList = 33,
        SummaryAct = 34,
        TransportAct = 35,
        PochtaAct = 36,
        Akt2Reestr = 37,
        Akt2ReestrPerso = 38,
        OfficeNote1 = 39,
        AccompDoc = 40,
        Courier7777 = 41,
        Book124Label = 42
    }
    public class MyProd
    {
        public int ID;
        public string Name;
        public string Bank;
        public int Type;
        public static int cntMax = 15;
        // счетчики для распоряжения на движение ценностей;
        public int[] cnts;
        public MyProd(int i, string name, string bank, int type)
        {
            ID = i; Name = name; Bank = bank; Type = type;
            cnts = new int[cntMax];
            for (int t = 0; t < cntMax; t++)
                cnts[t] = 0;
        }
        public MyProd(int i, string name, string bank)
        {
            ID = i; Name = name; Bank = bank; Type = -1;
            cnts = new int[cntMax];
            for (int t = 0; t < cntMax; t++)
                cnts[t] = 0;
        }
    }

    public class FuncClass
    {
        public static string ConnectionString => System.Configuration.ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;

        public static int GetFieldIndex(string dataFieldName, GridView gv)
        {
            for (int i = 0; i < gv.Columns.Count; i++)
            {
                BoundField bf = gv.Columns[i] as BoundField;
                if (bf != null && bf.DataField.ToLower() == dataFieldName.ToLower())
                    return i;
            }

            return -1;
        }
        public static int GetSortIndex(string SortFieldName, GridView gv)
        {
            for (int i = 0; i < gv.Columns.Count; i++)
            {
                BoundField bf = gv.Columns[i] as BoundField;
                TemplateField tf = gv.Columns[i] as TemplateField;
                if ((bf != null && bf.SortExpression.ToLower() == SortFieldName.ToLower()) || tf != null && tf.SortExpression.ToLower() == SortFieldName.ToLower())
                    return i;
            }

            return -1;
        }

        public static void HideFields(string s_fld, GridView gv)
        {
            string[] ar_fld = s_fld.Split(',');
            for (int i = 0; i < ar_fld.Length; i++)
                gv.Columns[GetFieldIndex(ar_fld[i], gv)].Visible = false;
        }
        // public static object LockObject = new object();
        public static string Bin2AHex(byte[] bytes)
        {
            if (bytes == null)
                return "";
            string str = "";
            foreach (byte b in bytes)
                str = $"{str}{b:X2}";
            return str;
        }

        public static string GetHashPan(string pan)
        {
            SHA1Managed sha = new SHA1Managed();
            string hashpan = Bin2AHex(sha.ComputeHash(Encoding.ASCII.GetBytes(pan)));
            return hashpan;
        }
        public static ClientType ClientType
        {
            get
            {
                string str = System.Configuration.ConfigurationManager.AppSettings["ClientType"];
                if (str == "Uzcard")
                    return ClientType.Uzcard;
                return ClientType.AkBars;
            }
        }

        public static TypeDoc GetTypeDoc(int iTypeDoc)
        {
            switch (iTypeDoc)
            {
                case ((int)TypeDoc.CardFromStorage): return TypeDoc.CardFromStorage;
                case ((int)TypeDoc.PinFromStorage): return TypeDoc.PinFromStorage;
                case ((int)TypeDoc.CardToStorage): return TypeDoc.CardToStorage;
                case ((int)TypeDoc.PinToStorage): return TypeDoc.PinToStorage;
                case ((int)TypeDoc.SendToFilial): return TypeDoc.SendToFilial;
                case ((int)TypeDoc.ReceiveToFilial): return TypeDoc.ReceiveToFilial;
                case ((int)TypeDoc.SendToClient): return TypeDoc.SendToClient;
                case ((int)TypeDoc.PersoCard): return TypeDoc.PersoCard;
                case ((int)TypeDoc.SendToBank): return TypeDoc.SendToBank;
                case ((int)TypeDoc.ReceiveToBank): return TypeDoc.ReceiveToBank;
                case ((int)TypeDoc.DeleteBrak): return TypeDoc.DeleteBrak;
                case ((int)TypeDoc.ReturnToFilial): return TypeDoc.ReturnToFilial;
                case ((int)TypeDoc.FilialFilial): return TypeDoc.FilialFilial;
                case ((int)TypeDoc.Expendables): return TypeDoc.Expendables;
                case ((int)TypeDoc.Reklama): return TypeDoc.Reklama;
                case ((int)TypeDoc.ToWrapping): return TypeDoc.ToWrapping;
                case ((int)TypeDoc.FromWrapping): return TypeDoc.FromWrapping;
                case ((int)TypeDoc.DontReceiveToFilial): return TypeDoc.DontReceiveToFilial;
                case ((int)TypeDoc.SendToFilialFilial): return TypeDoc.SendToFilialFilial;
                case ((int)TypeDoc.ReceiveToFilialExpertiza): return TypeDoc.ReceiveToFilialExpertiza;
                case ((int)TypeDoc.SendToExpertiza): return TypeDoc.SendToExpertiza;
                case ((int)TypeDoc.ReceiveToExpertiza): return TypeDoc.ReceiveToExpertiza;
                case ((int)TypeDoc.Expertiza): return TypeDoc.Expertiza;
                case ((int)TypeDoc.SendToPodotchet): return TypeDoc.SendToPodotchet;
                case ((int)TypeDoc.ReceiveToPodotchet): return TypeDoc.ReceiveToPodotchet;
                case ((int)TypeDoc.SendToClientFromPodotchet): return TypeDoc.SendToClientFromPodotchet;
                case ((int)TypeDoc.ReturnFromPodotchet): return TypeDoc.ReturnFromPodotchet;
                case ((int)TypeDoc.ReceiveFromPodotchet): return TypeDoc.ReceiveFromPodotchet;
                case ((int)TypeDoc.WriteOfPodotchet): return TypeDoc.WriteOfPodotchet;
                case ((int)TypeDoc.KillingCard): return TypeDoc.KillingCard;
                // ограничить их 30, иначе пойдут пересечения с TypeFormDoc (вот так там по идиотски сделано :(((
                case ((int)TypeDoc.ToBook124): return TypeDoc.ToBook124;
                case ((int)TypeDoc.GetBook124): return TypeDoc.GetBook124;
                case ((int)TypeDoc.FromBook124): return TypeDoc.FromBook124;
                case ((int)TypeDoc.ReceiveBook124): return TypeDoc.ReceiveBook124;
                case ((int)TypeDoc.SendToClientService): return TypeDoc.SendToClientService;
                case ((int)TypeDoc.ReceiveToFilialPacket): return TypeDoc.ReceiveToFilialPacket;
                case ((int)TypeDoc.ToGoz): return TypeDoc.ToGoz;
                case ((int)TypeDoc.GetGoz): return TypeDoc.GetGoz;
                case ((int)TypeDoc.FromGoz): return TypeDoc.FromGoz;
                case ((int)TypeDoc.ReceiveGoz): return TypeDoc.ReceiveGoz;
                case ((int)TypeDoc.FromGozToPodotchet): return TypeDoc.FromGozToPodotchet;
                case ((int)TypeDoc.ToPodotchetFromGoz): return TypeDoc.ToPodotchetFromGoz;
                case ((int)TypeDoc.FromPodotchetToGoz): return TypeDoc.FromPodotchetToGoz;
                case ((int)TypeDoc.ToGozFromPodotchet): return TypeDoc.ToGozFromPodotchet;
                default:
                    throw new Exception("Ошибка определения типа документа: неизвестный тип (см. FuncClass.GetTypeDoc)");
            }
        }


        public static CardStatus CardStatusInDoc_Confirmed(int typeDoc)
        {
            return CardStatusInDoc_Confirmed(GetTypeDoc(typeDoc));
        }
        public static CardStatus CardStatusInDoc_Confirmed(TypeDoc typeDoc)
        {
            switch (typeDoc)
            {
                case TypeDoc.ToBook124: return CardStatus.FilialBook124;
                case TypeDoc.GetBook124: return CardStatus.Book124;
                case TypeDoc.FromBook124: return CardStatus.Book124Filial;
                case TypeDoc.ReceiveBook124: return CardStatus.Filial;
            }
            return CardStatus.Undefined;
        }

        public static CardStatus CardStatusInDoc_NotConfirmed(int typeDoc)
        {
            return CardStatusInDoc_NotConfirmed(GetTypeDoc(typeDoc));
        }
        public static CardStatus CardStatusInDoc_NotConfirmed(TypeDoc typeDoc)
        {
            switch (typeDoc)
            {
                case TypeDoc.ToBook124: return CardStatus.Filial;
                case TypeDoc.GetBook124: return CardStatus.FilialBook124;
                case TypeDoc.FromBook124: return CardStatus.Book124;
                case TypeDoc.ReceiveBook124: return CardStatus.Book124Filial;
                case TypeDoc.ToGoz: return CardStatus.Filial;
                case TypeDoc.GetGoz: return CardStatus.FilialGoz;
                case TypeDoc.FromGoz: return CardStatus.Goz;
                case TypeDoc.ReceiveGoz: return CardStatus.GozFilial;
                case TypeDoc.FromGozToPodotchet: return CardStatus.Goz;
                case TypeDoc.ToGozFromPodotchet: return CardStatus.Accountable;
            }
            return CardStatus.Undefined;
        }

    }
    public class RMKData
    {
        public int iddoc;
        public int baseproducttype;
        public string guidrmk;
    }


    public class BranchStore
    {
        public int id;
        public string ident_dep;
        public string department;
        //в последнем индексе будет сумма предыдущих (используется, например, в возврате ценностей, когда надо посчитать и бракованные и неверные филиалы отправки)
        public int[] countMasterCard = new int[20];
        public int[] countVisaCard = new int[20];
        public int[] countServiceCard = new int[20];
        public int[] countPinConvert = new int[20];
        public int[] countNFCCard = new int[20];
        public int[] countMirCard = new int[20];

        
        public int countNonEmpty()
        {
            int count = 0;
            for (int i = 0; i < countMasterCard.Length; i++) if (countMasterCard[i] > 0) count++;
            for (int i = 0; i < countVisaCard.Length; i++) if (countVisaCard[i] > 0) count++;
            for (int i = 0; i < countServiceCard.Length; i++) if (countServiceCard[i] > 0) count++;
            for (int i = 0; i < countPinConvert.Length; i++) if (countPinConvert[i] > 0) count++;
            for (int i = 0; i < countNFCCard.Length; i++) if (countNFCCard[i] > 0) count++;
            for (int i = 0; i < countMirCard.Length; i++) if (countMirCard[i] > 0) count++;
            return count;
        }

        public bool isEmpty()
        {
            if (isEmptyMasterCard() && isEmptyVisaCard() && isEmptyServiceCard() && isEmptyPinConvert() && isEmptyNFCCard() && isEmptyMirCard()) return true;
            return false;
        }

        public bool isEmptyMasterCard()
        {
            if (countMasterCard[0] < 1 && countMasterCard[1] < 1 && countMasterCard[2] < 1 && countMasterCard[3] < 1) return true;
            return false;
        }

        public bool isEmptyVisaCard()
        {
            if(countVisaCard[0] < 1 && countVisaCard[1] < 1 && countVisaCard[2] < 1 && countVisaCard[3] < 1) return true;
            return false;
        }

        public bool isEmptyServiceCard()
        {
            if (countServiceCard[0] < 1 && countServiceCard[1] < 1 && countServiceCard[2] < 1 && countServiceCard[3] < 1) return true;
            return false;
        }

        public bool isEmptyPinConvert()
        {
            if(countPinConvert[0] < 1 && countPinConvert[1] < 1 && countPinConvert[2] < 1 && countPinConvert[3] < 1) return true;
            return false;
        }

        public bool isEmptyNFCCard()
        {
            if(countNFCCard[0] < 1 && countNFCCard[1] < 1 && countNFCCard[2] < 1 && countNFCCard[3] < 1) return true;
            return false;
        }

        public bool isEmptyMirCard()
        {
            if (countMirCard[0] < 1 && countMirCard[1] < 1 && countMirCard[2] < 1 && countMirCard[3] < 1) return true;
            return false;
        }

        public BranchStore(int id, string ident_dep, string department)
        {
            this.id = id;
            this.ident_dep = (ident_dep == null) ? "" : ident_dep;
            this.department = (department == null) ? "" : department;
        }
        public BranchStore(Object id, Object ident_dep, Object department)
        {
            this.id = Convert.ToInt32(id);
            this.ident_dep = (ident_dep == null) ? "" : Convert.ToString(ident_dep);
            this.department = (department == null) ? "" : Convert.ToString(department);
        }

        public void Clear()
        {
            for (int i = 0; i < countMasterCard.Length; i++) countMasterCard[i] = 0;
            for (int i = 0; i < countVisaCard.Length; i++) countVisaCard[i] = 0;
            for (int i = 0; i < countServiceCard.Length; i++) countServiceCard[i] = 0;
            for (int i = 0; i < countPinConvert.Length; i++) countPinConvert[i] = 0;
            for (int i = 0; i < countNFCCard.Length; i++) countNFCCard[i] = 0;
            for (int i = 0; i < countMirCard.Length; i++) countMirCard[i] = 0;
        }

        public int addCount(int type, string prodName, int indexCount, int count)
        {
            if (indexCount < 0 || indexCount > 19) return -1;
            BaseProductType pt=BranchStore.codeFromTypeAndProdName(type, prodName);
            switch (pt)
            {
                case BaseProductType.MasterCard:
                                                countMasterCard[indexCount] += count;
                                                break;
                case BaseProductType.VisaCard:
                                                countVisaCard[indexCount] += count;
                                                break;
                case BaseProductType.NFCCard:
                                                countNFCCard[indexCount] += count;
                                                break;
                case BaseProductType.ServiceCard:
                                                countServiceCard[indexCount] += count;
                                                break;
                case BaseProductType.MirCard:
                                                countMirCard[indexCount] += count;
                                                break;
                case BaseProductType.PinConvert:
                                                countPinConvert[indexCount] += count;
                                                break;
                default:
                                                break;
                                                
            }
            return (int)pt;
            /*
            if (type == 2) { countPinConvert[indexCount] += count; return BranchStore.PINCODE; }
            if (type == 1)
            {

                int code = BranchStore.codeFromTypeAndProdName(type, prodName);
                if (code==1)
                {
                    countMasterCard[indexCount] += count;
                    return 1;
                }
                else
                if (code==2)
                {
                    countVisaCard[indexCount] += count;
                    return 2;
                }
                else
                if (code==3)
                {
                    countNFCCard[indexCount] += count;
                    return 3;
                }
                if (code == 5)
                {
                    countMirCard[indexCount] += count;
                    return 5;
                }
                else
                {
                    countServiceCard[indexCount] += count;
                    return 4;
                }
            }
            return -2;
            */ 
        }

        static public BaseProductType codeFromTypeAndProdName(int type,String prodName)
        {
            if(type==1)
            {
                string prod_name = prodName.ToLower();
                if (prod_name.ToLower().StartsWith("mc") == true || prod_name.ToLower().StartsWith("master") || prod_name.ToLower().StartsWith("mifare standart"))
                {
                    return BaseProductType.MasterCard;
                }
                else
                if (prod_name.ToLower().StartsWith("visa"))
                {
                    return BaseProductType.VisaCard;
                }
                else
                if (prod_name.ToLower().StartsWith("nfs") == true || prod_name.ToLower().StartsWith("nfc"))
                {
                    return BaseProductType.NFCCard;
                }
                if (prod_name.ToLower().StartsWith("mir"))
                {
                    return BaseProductType.MirCard;
                }
                else
                {
                    return BaseProductType.ServiceCard;
                }
            }
            if (type == 2) return BaseProductType.PinConvert;
            else return BaseProductType.None;
        }

        public string getNameForCard(int code)
        {
            if (code == (int)BaseProductType.MasterCard) return "MasterCard (ИТОГО)";
            if (code == (int)BaseProductType.VisaCard) return "Visa (ИТОГО)";
            if (code == (int)BaseProductType.NFCCard) return "NFC карты (ИТОГО)";
            if (code == (int)BaseProductType.ServiceCard) return "Сервисные карты (ИТОГО)";
            if (code == (int)BaseProductType.MirCard) return "Карты МИР(ИТОГО)";
            return "";
        }

        public BranchStore getBaseProductsFromDocs(int id_doc, SqlConnection conn, SqlTransaction trans = null)
        {
            SqlCommand cmdSelect = new SqlCommand(string.Format("select Products_StorageDocs.*,Products.id_type,Products.name from Products_StorageDocs " +
                                                                "join Products_Banks on Products_Banks.id =  Products_StorageDocs.id_prb " +
                                                                "join Products on Products.id = Products_Banks.id_prod " +
                                                                "where id_doc={0}", id_doc), conn);
            BranchStore bs = this;
            bs.Clear();
            bool isProduct = false;
            cmdSelect.Transaction = trans;
            using (SqlDataReader dr = cmdSelect.ExecuteReader())
            {
                if (dr.HasRows)
                {
                    isProduct = true;
                    while (dr.Read() == true)
                    {
                        bs.addCount(Convert.ToInt32(dr["id_type"]), dr["name"].ToString(), 0, Convert.ToInt32(dr["cnt_new"]));
                        bs.addCount(Convert.ToInt32(dr["id_type"]), dr["name"].ToString(), 1, Convert.ToInt32(dr["cnt_perso"]));
                        bs.addCount(Convert.ToInt32(dr["id_type"]), dr["name"].ToString(), 2, Convert.ToInt32(dr["cnt_brak"]));
                        bs.addCount(Convert.ToInt32(dr["id_type"]), dr["name"].ToString(), 19, Convert.ToInt32(dr["cnt_new"]) + Convert.ToInt32(dr["cnt_perso"]) + Convert.ToInt32(dr["cnt_brak"]));
                    }
                }
                dr.Close();
            }
            if (isProduct == false)
            {
                cmdSelect = new SqlCommand(string.Format("select products.id_type,products.name,cards.ispin from cards_storagedocs " +
                                                "join cards on cards.id=cards_storagedocs.id_card " +
                                                "join products_banks on products_banks.id = cards.id_prb " +
                                                "join products on products.id=products_banks.id_prod " +
                                                "where cards_storagedocs.id_doc={0}", id_doc), conn);
                cmdSelect.Transaction = trans;
                using (SqlDataReader dr = cmdSelect.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        while (dr.Read() == true)
                        {
                            bs.addCount(Convert.ToInt32(dr["id_type"]), dr["name"].ToString(), 0, 1);
                            bs.addCount(Convert.ToInt32(dr["id_type"]), dr["name"].ToString(), 1, 1);
                            bs.addCount(Convert.ToInt32(dr["id_type"]), dr["name"].ToString(), 2, 1);
                            //if (Convert.ToInt32(dr["id_type"]) != 0)
                            if (Convert.ToInt32(dr["ispin"]) != 0)
                            {
                                bs.addCount(2, "PIN", 0, 1);
                                bs.addCount(2, "PIN", 1, 1);
                                bs.addCount(2, "PIN", 2, 1);
                                bs.addCount(2, "PIN", 19, 1);
                            }
                        }
                    }
                    dr.Close();
                }
            }
            return bs;
        }

        static public void clearRMKData(int id_doc, SqlConnection conn, SqlTransaction trans = null)
        {
            SqlCommand cmdDelete = new SqlCommand(string.Format("delete from storagedocsrmk where iddoc = {0}", id_doc), conn);
            cmdDelete.Transaction = trans;
            cmdDelete.ExecuteNonQuery();
        }

        static public void addRMKData(int id_doc, int baseproducttype, string guidrmk, SqlConnection conn, SqlTransaction trans = null)
        {
            SqlCommand cmdInsert = new SqlCommand(string.Format("insert into storagedocsrmk (iddoc, productbasetype, guidrmk) values({0},{1},'{2}')", id_doc, baseproducttype, guidrmk), conn);
            cmdInsert.Transaction = trans;
            cmdInsert.ExecuteNonQuery();
        }

        static public List<RMKData> getRMKData(int id_doc, SqlConnection conn, SqlTransaction trans = null)
        {
            SqlCommand cmdSelect = new SqlCommand(string.Format("select * from storagedocsrmk where iddoc={0}", id_doc), conn);
            cmdSelect.Transaction = trans;
            List<RMKData> rmkList = new List<RMKData>();
            using (SqlDataReader dr = cmdSelect.ExecuteReader())
            {
                if (dr.HasRows)
                {
                    while (dr.Read() == true)
                    {
                        RMKData r = new RMKData();
                        r.iddoc = Convert.ToInt32(dr["iddoc"]);
                        r.baseproducttype = Convert.ToInt32(dr["productbasetype"]);
                        r.guidrmk = dr["guidrmk"].ToString();
                        rmkList.Add(r);
                    }
                }
            }
            return rmkList;
        }

        static public string baseProductTypeToGuid(BaseProductType bpt)
        {
            string guid = "00000000-0000-0000-0000-000000000000";
            switch (bpt)
            {
                case BaseProductType.MasterCard: guid = "44ff86c8-b80c-46f7-a0e3-215676ee3e23"; break;
                case BaseProductType.VisaCard: guid = "191bf216-d26c-4b4d-9268-bb21db3315b0"; break;
                case BaseProductType.ServiceCard: guid = "a3bdce2b-910e-44f2-916e-a72bddfd0abd"; break;
                case BaseProductType.NFCCard: guid = "51b1f5fc-031d-4bdf-9a63-c58ba9de2b48"; break;
                case BaseProductType.MirCard: guid = "29ec811c-dc2c-4fb3-bf82-b1b0296f7d34"; break;
                case BaseProductType.PinConvert: guid = "61b96315-e3b4-4220-9fa3-3d3f9cde7c27"; break;
            }
            return guid;
        }

        static public int TypeDocToRMKOperation(TypeDoc td)
        {
            int typeop = 0;
            switch (td)
            {
                case TypeDoc.ReceiveToFilial: typeop = 1; break;
                case TypeDoc.ReceiveToFilialPacket: typeop = 1; break;
                case TypeDoc.DontReceiveToFilial: typeop = 2; break;
                case TypeDoc.FilialFilial: typeop = 3; break;
                case TypeDoc.SendToFilialFilial: typeop = 3; break; // 10.12.2019 добавили и для Передача ДО и ОК тип 3 как и Филиал-филиал
                case TypeDoc.SendToPodotchet: typeop = 4; break;
                case TypeDoc.SendToClient: typeop = 5; break;
                case TypeDoc.SendToClientService: typeop = 5; break;
                case TypeDoc.SendToClientFromPodotchet: typeop = 6; break;
                //case TypeDoc.ReturnFromPodotchet: typeop = 7; break; //19.12.2019 эту операцию переносим в ReceiveFromPodotchet
                case TypeDoc.ReceiveFromPodotchet: typeop = 7; break;
                case TypeDoc.SendToBank: typeop = 20; break; //21.05.2020 меняем 8 на 20
                //case TypeDoc.ReceiveToBank: typeop = 9; break;
                //case TypeDoc.DeleteBrak: typeop = 11; break;
                case TypeDoc.KillingCard: typeop = 11; break; // см выше, сперва 11 коду был сопоставлен Уничтожение ценностей в центре, 25.09.19 по просьбе Рустема стали сообщать в РМК об уничтожении карт в филиале и использовали тот же код
                //case TypeDoc.ReceiveToFilialExpertiza: typeop = 12; break;
                //case TypeDoc.SendToExpertiza: typeop = 13; break;                
                case TypeDoc.ToBook124: typeop = 14; break;
                case TypeDoc.ReceiveBook124: typeop = 15; break;

                case TypeDoc.ToGoz: typeop = 21; break;
                case TypeDoc.ReceiveGoz: typeop = 22; break;
                case TypeDoc.FromGozToPodotchet: typeop = 24; break;
                case TypeDoc.ToGozFromPodotchet: typeop = 26; break;
            }
            return typeop;
        }
        static public int TypeDocToRMKOperation(TypeDoc td, int tp)
        {
            if (td == TypeDoc.FilialFilial && tp == 1)
                return 8;
            if (td == TypeDoc.ReceiveToFilial && tp == 1)
                return 9;
            return TypeDocToRMKOperation(td);
        }
        static public int TypeDocToIndexBranchStore(TypeDoc td)
        {
            int indx = 0;
            switch (td)
            {
                case TypeDoc.ReceiveToFilial: indx = 1; break;
                case TypeDoc.ReceiveToFilialPacket: indx = 1; break;
                case TypeDoc.ReturnToFilial: indx = 1; break;
                case TypeDoc.DontReceiveToFilial: indx = 2; break;
                case TypeDoc.FilialFilial: indx = 1; break;
                case TypeDoc.SendToClient: indx = 1; break;
                case TypeDoc.SendToBank: indx = 19; break; //сумма бракованных и хороших (на пересылку в другие филиалы)
                case TypeDoc.ReceiveToBank: indx = 2; break;
                case TypeDoc.DeleteBrak: indx = 2; break;
                case TypeDoc.ReceiveToFilialExpertiza: indx = 1; break;
                case TypeDoc.SendToExpertiza: indx = 1; break;
                case TypeDoc.SendToClientFromPodotchet: indx = 1; break;
                case TypeDoc.ReturnFromPodotchet: indx = 1; break;
                case TypeDoc.ToBook124: indx = 1; break;
                case TypeDoc.ReceiveBook124: indx = 1; break;
                case TypeDoc.ToGoz: indx = 1; break;
                case TypeDoc.ReceiveGoz: indx = 1; break;
                case TypeDoc.FromGozToPodotchet: indx = 1; break;
                case TypeDoc.ToGozFromPodotchet: indx = 1; break;
            }
            return indx;
        }

        //----------------------------------------------------
        // Return value:
        // 0  - AkBarsAdmin
        // -1 - подразделения для возврата в головной офис
        // >0 - id филиала, у которого branchId подчиненный
        //----------------------------------------------------
        static public int getBranchMainFilial(int branchId,bool perso)
        {
            int branch_main_filial = 0;
            lock (Database.lockObjectDB)
            {
                if (branchId != 0 && perso == false)
                {
                    branch_main_filial = -1;
                    object obj = null;
                    Database.ExecuteScalar(String.Format("select top 1 case when ((select id_parent from Branchs where  id={0}) !=0) then (select id_parent from Branchs where  id={0}) else {0} end as id from Branchs", branchId), ref obj, null);
                    if (obj != null)
                    {
                        /*if (Convert.ToInt32(obj) == 106)*/ // убрать этот if, когда будут филиалы кроме казанского
                            branch_main_filial = Convert.ToInt32(obj);
                    }
                }
            }
            return branch_main_filial;
        }

        public string RurPhrase(decimal money)
		{
    		return CurPhrase(money, "рубль", "рубля", "рублей", true, "копейка", "копейки", "копеек", false);
		}

		public string UsdPhrase(decimal money)
		{
    		return CurPhrase(money, "доллар США", "доллара США", "долларов США", true, "цент", "цента", "центов", true);
		}

		public string NumPhrase(ulong Value, bool IsMale)
		{
		if (Value == 0UL) return "Ноль";
		string[] Dek1 = { "", " од", " дв", " три", " четыре", " пять", " шесть", " семь", " восемь", " девять", " десять", " одиннадцать", " двенадцать", " тринадцать", " четырнадцать", " пятнадцать", " шестнадцать", " семнадцать", " восемнадцать", " девятнадцать" };
		string[] Dek2 = { "", "", " двадцать", " тридцать", " сорок", " пятьдесят", " шестьдесят", " семьдесят", " восемьдесят", " девяносто" };
		string[] Dek3 = { "", " сто", " двести", " триста", " четыреста", " пятьсот", " шестьсот", " семьсот", " восемьсот", " девятьсот" };
		string[] Th = { "", "", " тысяч", " миллион", " миллиард", " триллион", " квадрилион", " квинтилион" };
		string str = "";
		for (byte th = 1; Value > 0; th++)
			{
			ushort gr = (ushort)(Value % 1000);
			Value = (Value - gr) / 1000;
			if (gr > 0)
				{
				byte d3 = (byte)((gr - gr % 100) / 100);
				byte d1 = (byte)(gr % 10);
				byte d2 = (byte)((gr - d3 * 100 - d1) / 10);
				if (d2 == 1) d1 += (byte)10;
				bool ismale = (th > 2) || ((th == 1) && IsMale);
				str = Dek3[d3] + Dek2[d2] + Dek1[d1] + EndDek1(d1, ismale) + Th[th] + EndTh(th, d1) + str;
				};
			};
		str = str.Substring(1, 1).ToUpper() + str.Substring(2);
		return str;
		}

		private string CurPhrase(decimal money,string word1, string word234, string wordmore, bool IsMale,string sword1, string sword234, string swordmore, bool sIsMale)
		{
		money = decimal.Round(money, 2);
		decimal decintpart = decimal.Truncate(money);
		ulong intpart = decimal.ToUInt64(decintpart);
		string str = NumPhrase(intpart, IsMale) + " ";
		byte endpart = (byte)(intpart % 100UL);
		if (endpart > 19) endpart = (byte)(endpart % 10);
		byte fracpart = decimal.ToByte((money - decintpart) * 100M);
		//str += "и " + ((fracpart < 10) ? "0" : "") + fracpart.ToString() + "/100 ";
		switch (endpart)
			{
			case 1: str += word1; break;
			case 2:
			case 3:
			case 4: str += word234; break;
			default: str += wordmore; break;
			}
		return str;
		}

		private static string EndTh(byte ThNum, byte Dek)
		{
		bool In234 = ((Dek >= 2) && (Dek <= 4));
		bool More4 = ((Dek > 4) || (Dek == 0));
		if (((ThNum > 2) && In234) || ((ThNum == 2) && (Dek == 1))) return "а";
		else if ((ThNum > 2) && More4) return "ов";
		else if ((ThNum == 2) && In234) return "и";
		else return "";
		}

		private static string EndDek1(byte Dek, bool IsMale)
		{
		if ((Dek > 2) || (Dek == 0)) return "";
		else if (Dek == 1)
			{
			if (IsMale) return "ин";
			else return "на";
			}
		else
			{
			if (IsMale) return "а";
			else return "е";
			}
		}

	}
    public enum BaseProductType
    {
        None=0,
        MasterCard=1,
        VisaCard=2,
        NFCCard=3,
        ServiceCard=4,
        PinConvert=5,
        MirCard=6
    }
    public enum AccountBranchType
    {
        None=0,
        In=1,
        Out=2,
        Return=3
    }

    public class AccountBranch
    {
        public BaseProductType productType;
        public AccountBranchType accountType;
        public String accountDebet;
        public String accountCredit;

        public AccountBranch()
        {
            productType = BaseProductType.None;
            accountType = AccountBranchType.None;
        }
    }

    public class OperationDay
    {
        public String operdays;
        public String operdaye;
        public bool isTomorrow;
        public bool parent;
        public bool isShift; // признак что задана смена
        public String shift1s;
        public String shift1e;
        public String shift2s;
        public String shift2e;
        private DataSet ds = new DataSet();

        static public int getBranchSeqNumber(int branchId)
        {
            
            //String sql = "SELECT rownum from Branchs" +
            //           " JOIN ( SELECT ROW_NUMBER() over (order by id) as rownum,id as idbranch FROM Branchs where id=106 or id_parent=106) as b on b.idbranch=Branchs.id" +
            //           " WHERE id=" + branchId.ToString();
            String sql = "SELECT rownum from Branchs" +
                       " JOIN ( SELECT ROW_NUMBER() over (order by id) as rownum,id as idbranch FROM Branchs) as b on b.idbranch=Branchs.id" +
                       " WHERE id=" + branchId.ToString();
            object obj = null;
            lock (Database.lockObjectDB)
            {
                Database.ExecuteScalar(sql, ref obj, null);
                if (obj != null && obj != DBNull.Value)
                {
                    return Convert.ToInt32(obj);
                }
                return -1;
            }
        }

        static public int getShiftNumber()
        {
            return 20;
        }

        static public int getBranchStartNumber(int branchId)
        {
            int seqNum = getBranchSeqNumber(branchId);
            
            if (seqNum > 0)
            {
                seqNum = (seqNum - 1) * getShiftNumber()*2 + 1;
            }
            return seqNum;
        }

        static public String getBranchStartNumber(int branchId,DateTime curDate,bool first)
        {
            return (curDate.DayOfYear.ToString().PadLeft(3,'0') + String.Format("{0:yy}",curDate) + branchId.ToString().PadLeft(4,'0') + ((first==true) ? "1":"2"));
        }

        static public String getBranchStartNumber(int branchId, DateTime curDate, int num, int smena)
        {
            String n = ((curDate.DayOfYear + 100) * 1000 + branchId * 100 + num * 10 + smena).ToString();
            if (n.Length > 6) return n.Substring(0, 6);
            return n;
        }
       
        public OperationDay()
        {
            operdays=null;
            operdaye=null;
            isTomorrow=true;
            parent = true;
            isShift = false;
            shift1s = null;
            shift1e = null;
            shift2s = null;
            shift2e = null;
        }


        public void read(int branchId)
        {
            operdays=null;
            operdaye=null;
            isTomorrow=true;
            parent = true;
            isShift = false;
            shift1s = null;
            shift1e = null;
            shift2s = null;
            shift2e = null;
            lock (Database.lockObjectDB)
            {
                ds.Clear();
                Database.ExecuteQuery(String.Format("select isshift,shift1s,shift1e,shift2s,shift2e from Branchs where id={0}", branchId), ref ds, null);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    isShift = Convert.ToBoolean(ds.Tables[0].Rows[0]["isshift"]);
                    if (isShift == true)
                    {
                        String S1 = (ds.Tables[0].Rows[0]["shift1s"] != DBNull.Value) ? ds.Tables[0].Rows[0]["shift1s"].ToString() : "";
                        String E1 = (ds.Tables[0].Rows[0]["shift1e"] != DBNull.Value) ? ds.Tables[0].Rows[0]["shift1e"].ToString() : "";
                        String S2 = (ds.Tables[0].Rows[0]["shift2s"] != DBNull.Value) ? ds.Tables[0].Rows[0]["shift2s"].ToString() : "";
                        String E2 = (ds.Tables[0].Rows[0]["shift2e"] != DBNull.Value) ? ds.Tables[0].Rows[0]["shift2e"].ToString() : "";
                        if (S1.Length > 0 && E1.Length > 0 && S2.Length > 0 && E2.Length > 0)
                        {
                            shift1s = S1;
                            shift1e = E1;
                            shift2s = S2;
                            shift2e = E2;
                        }
                        else isShift = false;
                    }
                }
                ds.Clear();
                Database.ExecuteQuery(String.Format("select id,id_parent,operdays,operdaye,isodetomorrow from Branchs where id={0} or (id_parent=0 and id=(select id_parent from Branchs where id={0}))", branchId), ref ds, null);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    for (int n = 0; n < 2; n++)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {

                            String S = (ds.Tables[0].Rows[i]["operdays"] != DBNull.Value) ? ds.Tables[0].Rows[i]["operdays"].ToString() : "";
                            String E = (ds.Tables[0].Rows[i]["operdaye"] != DBNull.Value) ? ds.Tables[0].Rows[i]["operdaye"].ToString() : "";
                            if (S.Length > 0 && E.Length > 0)
                            {
                                operdays = S;
                                operdaye = E;
                                if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_parent"]) > 0) parent = false;
                                isTomorrow = Convert.ToBoolean(ds.Tables[0].Rows[i]["isodetomorrow"]);
                                if (n == 0 && Convert.ToInt32(ds.Tables[0].Rows[i]["id"]) == branchId) return;
                            }
                        }
                    }
                }
            }
        }

        public void write(int branchId)
        {
            //if (parent == true) return;
            lock (Database.lockObjectDB)
            {
                SqlCommand sqCom = new SqlCommand();
                sqCom.CommandText = "update Branchs set operdays=@operdays,operdaye=@operdaye,isodetomorrow=@isodetomorrow,isshift=@isshift,shift1s=@shift1s,shift1e=@shift1e,shift2s=@shift2s,shift2e=@shift2e where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = branchId;
                operdays=operdays.Trim();
                operdaye=operdaye.Trim();
                if ((operdays == null && operdaye == null) || (operdays.Length < 1 && operdaye.Length < 1))
                {
                    sqCom.Parameters.Add("@isodetomorrow", SqlDbType.Bit).Value = true;
                    sqCom.Parameters.Add("@operdays", SqlDbType.VarChar, 8).Value = DBNull.Value;
                    sqCom.Parameters.Add("@operdaye", SqlDbType.VarChar, 8).Value = DBNull.Value;
                }
                else
                {
                    DateTime s, e;
                    try
                    {
                        s = DateTime.ParseExact(operdays, "HH:mm", null);
                    }
                    catch (Exception e1)
                    {
                        throw new Exception("Неверное время начала: " + e1.Message);
                    }
                    try
                    {
                        e = DateTime.ParseExact(operdaye, "HH:mm", null);
                    }
                    catch (Exception e2)
                    {
                        throw new Exception("Неверное время окончания: " + e2.Message);
                    }
                    if (isTomorrow == false && s.CompareTo(e) > 0)
                        throw new Exception("Время начала позже времени окончания");

                    DateTime dtS = getDateTimeStart(DateTime.Now);
                    DateTime dtE = getDateTimeEnd(DateTime.Now);
                    TimeSpan left = dtE - dtS;
                    double hleft = left.TotalHours;
                    if (hleft < 0.0)
                        throw new Exception("Время начала позже времени окончания");
                    if (hleft > 24.0)
                        throw new Exception("Разница во времени превышает 24 часа");
                    sqCom.Parameters.Add("@isodetomorrow", SqlDbType.Bit).Value = isTomorrow;
                    sqCom.Parameters.Add("@operdays", SqlDbType.VarChar, 8).Value = operdays;
                    sqCom.Parameters.Add("@operdaye", SqlDbType.VarChar, 8).Value = operdaye;
                }
                if (isShift == false)
                {
                    sqCom.Parameters.Add("@isshift", SqlDbType.Bit).Value = isShift;
                    sqCom.Parameters.Add("@shift1s", SqlDbType.VarChar, 8).Value = DBNull.Value;
                    sqCom.Parameters.Add("@shift1e", SqlDbType.VarChar, 8).Value = DBNull.Value;
                    sqCom.Parameters.Add("@shift2s", SqlDbType.VarChar, 8).Value = DBNull.Value;
                    sqCom.Parameters.Add("@shift2e", SqlDbType.VarChar, 8).Value = DBNull.Value;
                }
                else
                {
                    try
                    {
                        DateTime.ParseExact(shift1s, "HH:mm", null);
                        DateTime.ParseExact(shift1e, "HH:mm", null);
                        DateTime.ParseExact(shift2s, "HH:mm", null);
                        DateTime.ParseExact(shift2e, "HH:mm", null);
                    }
                    catch (Exception e1)
                    {
                        throw new Exception("Неверно заданы времена смены: " + e1.Message);
                    }
                    DateTime[] sh1 = getDateTimePart(DateTime.Now, true);
                    DateTime[] sh2 = getDateTimePart(DateTime.Now, false);
                    if (sh1[0] > sh1[1]) throw new Exception("Время начала первой смены больше времени окончания. " + getMessagePart());
                    if (sh2[0] > sh2[1]) throw new Exception("Время начала второй смены больше времени окончания. " + getMessagePart());
                    if (sh1[0] >= sh2[0] || sh1[1] > sh2[0] || sh2[0] < sh1[1] || sh2[1] < sh1[0])
                        throw new Exception("Пересечение смен не допустимо. " + getMessagePart());
                    DateTime dtS = getDateTimeStart(DateTime.Now);
                    DateTime dtE = getDateTimeEnd(DateTime.Now);
                    if (sh1[0] < dtS || sh2[1] > dtE) throw new Exception("Cмены выходят за границы опердня. " + getMessagePart());
                    
                    sqCom.Parameters.Add("@isshift", SqlDbType.Bit).Value = isShift;
                    sqCom.Parameters.Add("@shift1s", SqlDbType.VarChar, 8).Value = shift1s;
                    sqCom.Parameters.Add("@shift1e", SqlDbType.VarChar, 8).Value = shift1e;
                    sqCom.Parameters.Add("@shift2s", SqlDbType.VarChar, 8).Value = shift2s;
                    sqCom.Parameters.Add("@shift2e", SqlDbType.VarChar, 8).Value = shift2e;
                    
                }
                Database.ExecuteNonQuery(sqCom, null);
            }
        }
        public bool isParent()
        {
            return parent;
        }

        public bool isCurrentDay()
        {
            if ((operdays == null && operdaye == null) || (operdays.Length < 1 && operdaye.Length < 1)) return true;
            else return false;
        }

        public DateTime roundDateTime(DateTime curDate)
        {

            String dateStart = String.Format("{0:dd.MM.yyyy}", curDate);
            return DateTime.ParseExact(dateStart, "dd.MM.yyyy", null);
        }
        public DateTime getDateTimeStart(DateTime curDate)
        {

            String timeStart = "00:00";
            String dateStart = String.Format("{0:dd.MM.yyyy}", curDate);
            if (operdays != null && operdays.Length>0)
                dateStart += " " + operdays;
            else
             dateStart += " " + timeStart;
            DateTime ret = DateTime.ParseExact(dateStart, "dd.MM.yyyy HH:mm", null);
            if (isTomorrow == true && isCurrentDay()==false)
            {
                ret=ret.AddDays(-1); //-=TimeSpan.FromDays(1);
            }
            return ret;
            //return DateTime.ParseExact(dateStart, "dd.MM.yyyy HH:mm", null);
        }
        public DateTime getDateTimeEnd(DateTime curDate)
        {
            String timeEnd = "00:00";
            String dateEnd = String.Format("{0:dd.MM.yyyy}", curDate);
            if (operdaye != null && operdaye.Length > 0)
                dateEnd += " " + operdaye;
            else
                dateEnd += " " + timeEnd;
            DateTime ret = DateTime.ParseExact(dateEnd, "dd.MM.yyyy HH:mm", null);
            if (isCurrentDay()==true)
            {
                ret=ret.AddDays(1); //+=TimeSpan.FromDays(1);
            }
            return ret;
        }
        public String getMessage(DateTime curDate)
        {
            return String.Format("c {0:dd.MM.yyyy HH:mm} до {1:dd.MM.yyyy HH:mm}",getDateTimeStart(curDate),getDateTimeEnd(curDate));
        }

        public DateTime[] getDateTimePart(DateTime curDate, bool first)
        {
            if(isShift==false) return null;
            DateTime[] tm = new DateTime[2];
            DateTime tmFirst;
            DateTime tmNext;
            if (first == true)
            {
                DateTime s=DateTime.ParseExact(shift1s, "HH:mm", null);
                tmFirst = getDateTimeStart(curDate);
                tmFirst = new DateTime(tmFirst.Year, tmFirst.Month, tmFirst.Day, s.Hour, s.Minute, s.Second);
                DateTime e = DateTime.ParseExact(shift1e, "HH:mm", null);
                tmNext = getDateTimeStart(curDate);
                tmNext = new DateTime(tmNext.Year, tmNext.Month, tmNext.Day, e.Hour, e.Minute, e.Second);
                if(e<s && isTomorrow==true) tmNext=tmNext.AddDays(1); // += TimeSpan.FromDays(1);
                
            }
            else
            {
                DateTime e = DateTime.ParseExact(shift2e, "HH:mm", null);
                tmNext = getDateTimeEnd(curDate);
                tmNext = new DateTime(tmNext.Year, tmNext.Month, tmNext.Day, e.Hour, e.Minute, e.Second);
                DateTime s = DateTime.ParseExact(shift2s, "HH:mm", null);
                tmFirst = getDateTimeEnd(curDate);
                tmFirst = new DateTime(tmFirst.Year, tmFirst.Month, tmFirst.Day, s.Hour, s.Minute, s.Second);
                if (e < s && isTomorrow == true) tmFirst=tmFirst.AddDays(-1); // -= TimeSpan.FromDays(1);
            }
            tm[0] = tmFirst;
            tm[1] = tmNext;
            return tm;
        }

        public String getMessagePart(DateTime curDate, bool first)
        {
            DateTime[] tm = getDateTimePart(curDate, first);
            if (tm != null) return String.Format("c {0:dd.MM.yyyy HH:mm} до {1:dd.MM.yyyy HH:mm}", tm[0], tm[1]);
            else return "";
        }
        
        public String getMessagePart()
        {
            if (isShift == false) return "";
            try
            {
                DateTime[] sh1 = getDateTimePart(DateTime.Now, true);
                DateTime[] sh2 = getDateTimePart(DateTime.Now, false);
                return "Расчетные значения - I: " + getMessagePart(DateTime.Now, true) + ", II: " + getMessagePart(DateTime.Now, false);
            }
            catch (Exception)
            {
            }
            return "Расчетные значения - ошибка";
        }

        public bool isEmpty()
        {
            if (operdays == null || operdaye == null || operdays.Length < 1 || operdaye.Length < 1) return true;
            return false;
        }
    }

    public class SingleQuery
    {
        public static bool IsAccountable(string login)
        {
            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
            {
                conn.Open();
                using (SqlCommand comm = conn.CreateCommand())
                {
                    comm.CommandText = $"select count(*) from V_UserAction where UserName='{login}' and ActionId in ({(int)Restrictions.ReceiveToPodotchet}, {(int)Restrictions.SendToClientFromPodotchet}, {(int)Restrictions.ReturnFromPodotchet})";
                    object obj = 0;
                    try
                    {
                        obj = comm.ExecuteScalar();
                    }
                    catch (Exception e)
                    {
                        LogClass.WriteToLog(e.ToString());
                    }
                    conn.Close();
                    return ((int) obj > 0);
                }
            }
        }
        public static void ExecuteNonQuery(SqlCommand comm)
        {
            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
            {
                conn.Open();
                comm.Connection = conn;
                // exception ловить здесь не буду. Пусть ловится выше
                comm.ExecuteNonQuery();
                conn.Close();
            }
        }

        public static string BranchName(int branchId)
        {
            string res = "";
            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
            {
                conn.Open();
                using (SqlCommand comm = conn.CreateCommand())
                {
                    comm.CommandText = $"select ident_dep from Branchs where id={branchId}";
                    object obj = comm.ExecuteScalar();
                    res = (obj == DBNull.Value || obj == null) ? "" : obj.ToString();
                }
                conn.Close();
            }
            return res;
        }
    }

    public enum ClientType
    {
        AkBars,
        Uzcard
    }
}
