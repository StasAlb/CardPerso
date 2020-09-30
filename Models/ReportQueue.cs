using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace JQueryDataTables.Models
{
    public class Report
    {
        public int id { get; set; }
        public string Title
        {
            get
            {
                switch (Type)
                {
                    case 1: return "Записка на выпуск карт";
                    case 2: return "Cорт лист";
                    case 3: return "Выпущенные карты: по продуктам";
                    case 4: return "Выпущенные карты: динамика";
                    case 6: return "Хранилище: текущее состояние";
                    case 7: return "Хранилище: ценности с количеством ниже порогового значения";
                    case 8: return "Уничтоженные карты:по продуктам";
                    case 9: return "Уничтоженные карты: по подразделениям";
                    case 10: return "Бракованные карты: по продуктам";
                    case 11: return "Закупочные договора";
                    case 12: return "Сводный акт получениея ценностей";
                    case 14: return "Выданные ценности";
                    case 15: return "Движение ценностей";
                    case 16: return "Выданные карты по подразделению";
                    case 17: return "Выданные карты по филиалу";
                    case 18: return "Хранилище текущее состоянии_1";
                    case 19: case 41: return "Филиал итог день";
                    case 42: return "Отчет МЕМОРИАЛЬНЫЙ ОРДЕР";
                    case 20: return "Консолидированный отчет по Казанскому филиалу";
                    case 43: return "Списание карт/пин конвертов с подотчета МОЛ";
                    case 44: return "Карта жителя";
                    case 45: return "Отчет выданных карт по книге 124";
                    default: return "Неизвестный тип отчета";
                }
            }
        }
        public string Date { get; set; }
        public int Status { get; set; }
        public string StatusName
        {
            get
            {
                switch (Status)
                {
                    case 1:
                        return $"{WaitCnt+1} в очереди  <img width='36px' height='16px' src='/Images/Wait.gif'/>";
                    case 2:
                        return "Идет формирование";
                    case 3:
                        return "Готов";
                    case 666:
                    case 999:
                        return "Ошибка";
                    default:
                        return $"Ошибка ({Status})";
                }
            }
        }
        public int Type { get; set; }
        public string Parameters { get; set; }
        public string ButtonString
        {
            get
            {
                return "";
            }
        }
        public int WaitCnt { set; get; }
    }
}