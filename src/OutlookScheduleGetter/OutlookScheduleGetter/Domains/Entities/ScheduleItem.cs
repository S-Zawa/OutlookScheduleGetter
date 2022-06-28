using System;

namespace OutlookScheduleGetter.Domains.Entities
{
    /// <summary>
    /// Outlookから取得したスケジュールデータを管理
    /// </summary>
    public class ScheduleItem
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="title"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        public ScheduleItem(string title, DateTime start, DateTime end)
        {
            Title = title;
            Start = start;
            End = end;
        }

        /// <summary>
        /// タイトル
        /// </summary>
        public string Title { get; }

        /// <summary>
        /// 開始日時
        /// </summary>
        public DateTime Start { get; }

        /// <summary>
        /// 終了日時
        /// </summary>
        public DateTime End { get; }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"{Title}:{Start.ToString(AppSettings.DateTimeFormat)} - {End.ToString(AppSettings.DateTimeFormat)}";
        }
    }
}
