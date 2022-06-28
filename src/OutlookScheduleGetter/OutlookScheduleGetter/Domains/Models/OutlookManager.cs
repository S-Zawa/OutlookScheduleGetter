using System;
using System.Collections.Generic;
using System.Linq;
using NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using OutlookScheduleGetter.Domains.Entities;

namespace OutlookScheduleGetter.Domains.Models
{
    /// <summary>
    /// Outlookのスケジュールを取得する
    /// </summary>
    public class OutlookManager
    {
        /// <summary>
        /// スケジュール一覧取得
        /// </summary>
        /// <param name="mailOrName"></param>
        /// <param name="date"></param>
        /// <returns>スケジュール一覧</returns>
        public IEnumerable<ScheduleItem> GetScheduleList(string mailOrName, DateTime date)
        {
            return GetScheduleList(mailOrName, date, date.AddDays(1));
        }

        /// <summary>
        /// スケジュール一覧取得
        /// </summary>
        /// <param name="mailOrName"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns>スケジュール一覧</returns>
        public IEnumerable<ScheduleItem> GetScheduleList(string mailOrName, DateTime start, DateTime end)
        {
            var folder = GetCalenderFolder(mailOrName);
            var items = GetScheduleItems(folder, start, end);

            var schedules = ExpansionAndConvert(items, start, end);

            return schedules;
        }

        /// <summary>
        /// ScheduleItemに変換
        /// </summary>
        /// <param name="items"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns>ScheduleItem</returns>
        private IEnumerable<ScheduleItem> ExpansionAndConvert(_Items items, DateTime start, DateTime end)
        {
            var ret = new List<ScheduleItem>();

            foreach (AppointmentItem item in items)
            {
                if (item.IsRecurring)
                {
                    var pattern = item.GetRecurrencePattern();
                    DateTime first = new DateTime(start.Year, start.Month, start.Day, item.Start.Hour, item.Start.Minute, 0);
                    DateTime last = new DateTime(end.Year, end.Month, end.Day);
                    for (DateTime cur = first; cur <= last; cur = cur.AddDays(1))
                    {
                        try
                        {
                            // ここの取得が成功したらスケジュールが存在するって判定になる
                            var recur = pattern.GetOccurrence(cur);
                            var curEnd = new DateTime(cur.Year, cur.Month, cur.Day, item.End.Hour, item.End.Minute, 0);
                            ret.Add(new ScheduleItem(item.Subject, cur, curEnd));
                        }
                        catch
                        {
                        }
                    }
                }
                else
                {
                    ret.Add(new ScheduleItem(item.Subject, item.Start, item.End));
                }
            }

            ret = ret.Where(x => start <= x.Start)
                     .Where(x => x.End <= end)
                     .OrderBy(x => x.Start)
                     .ToList();

            return ret;
        }

        /// <summary>
        /// スケジュールの検索
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns>スケジュール</returns>
        private _Items GetScheduleItems(MAPIFolder folder, DateTime start, DateTime end)
        {
            string startDate = start.ToString("yy/MM/dd");
            string endDate = end.ToString("yy/MM/dd");
            string filter = $"[Start] >= '{startDate}' AND [Start] <= '{endDate}'";
            var list = folder.Items.Restrict(filter);

            return list;
        }

        /// <summary>
        /// 対象フォルダの取得
        /// </summary>
        /// <param name="mailOrName"></param>
        /// <returns>対象フォルダ</returns>
        private MAPIFolder GetCalenderFolder(string mailOrName)
        {
            var outlook = new Application();

            var recipient = outlook.Session.CreateRecipient(mailOrName);
            return outlook.Session.GetSharedDefaultFolder(recipient, OlDefaultFolders.olFolderCalendar);
        }
    }
}
