using System;

namespace OutlookScheduleGetter.Domains
{
    /// <summary>
    /// appsettings.json
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// 日付時刻フォーマット
        /// </summary>
        public static readonly string DateTimeFormat = "yyyy/MM/dd(ddd) HH:mm";

        /// <summary>
        /// MailOrName
        /// </summary>
        public string? MailOrName { get; set; }

        /// <summary>
        /// 開始日時
        /// </summary>
        public DateTime Start { get; set; }

        /// <summary>
        /// 終了日時
        /// </summary>
        public DateTime End { get; set; }
    }
}
