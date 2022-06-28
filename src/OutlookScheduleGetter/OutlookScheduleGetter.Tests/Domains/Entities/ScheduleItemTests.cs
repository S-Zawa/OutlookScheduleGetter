using System;
using OutlookScheduleGetter.Domains.Entities;
using Xunit;

namespace OutlookScheduleGetter.Tests.Domains.Entities
{
    public class ScheduleItemTests
    {
        [Fact]
        public void ToStringFormat()
        {
            var start = new DateTime(2022, 1, 1, 1, 11, 0);
            var end = start.AddDays(1);
            var scheduleItem = new ScheduleItem("テスト", start, end);
            var toString = $"テスト:2022/01/01(土) 01:11 - 2022/01/02(日) 01:11";
            Assert.Equal(toString, scheduleItem.ToString());
        }
    }
}
