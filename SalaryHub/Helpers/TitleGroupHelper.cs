namespace SalaryHub.Helpers
{
    public static class TitleGroupHelper
    {
        public const string GroupSummaryPrefix = "##tổng nhóm: ";

        public static int GetGroupOrder(string title)
        {
            var lower = title.ToLower();

            if (lower.Contains("kinh doanh bancas"))
                return 3;
            if (lower.Contains("kinh doanh"))
                return 2;
            if (lower.Contains("thu nhập"))
                return 1;

            return 4;
        }

        public static string GetGroupName(int group)
        {
            return group switch
            {
                1 => "thu nhập",
                2 => "kinh doanh",
                3 => "kinh doanh bancas",
                _ => "khác"
            };
        }

        public static bool IsGroupSummary(string title)
        {
            return title.StartsWith(GroupSummaryPrefix);
        }

        /// <summary>
        /// Không tính vào tổng thu nhập chịu thuế: cột tổng nhóm
        /// </summary>
        public static bool ShouldExcludeFromTotal(string title)
        {
            return IsGroupSummary(title);
        }

        public static string GetCssColor(string title)
        {
            if (IsGroupSummary(title))
            {
                // Lấy group order từ tên nhóm gốc
                var groupName = title.Substring(GroupSummaryPrefix.Length);
                return GetGroupOrderByName(groupName) switch
                {
                    1 => "#C6D9B3",
                    2 => "#B0CFE7",
                    3 => "#F0D4AD",
                    _ => "#C5BBD9"
                };
            }

            return GetGroupOrder(title) switch
            {
                1 => "#D9EAD3",
                2 => "#CFE2F3",
                3 => "#FCE5CD",
                _ => "#D9D2E9"
            };
        }

        public static string GetExcelColor(string title)
        {
            return GetCssColor(title);
        }

        public static List<string> SortByGroup(List<string> titles)
        {
            return titles
                .OrderBy(t => GetGroupOrder(t))
                .ToList();
        }

        /// <summary>
        /// Sắp xếp theo nhóm và chèn cột tổng nhóm cuối mỗi nhóm (nếu nhóm có >= 2 title)
        /// </summary>
        public static List<string> SortAndInsertGroupSummaries(List<string> titles)
        {
            var sorted = SortByGroup(titles);
            var result = new List<string>();

            var groups = sorted.GroupBy(t => GetGroupOrder(t)).OrderBy(g => g.Key);

            foreach (var group in groups)
            {
                var items = group.ToList();
                result.AddRange(items);

                if (items.Count >= 2)
                {
                    result.Add(GroupSummaryPrefix + GetGroupName(group.Key));
                }
            }

            return result;
        }

        public static Dictionary<string, string> BuildColorMap(List<string> titles)
        {
            return titles.ToDictionary(t => t, t => GetCssColor(t));
        }

        public static int GetGroupOrderByName(string groupName)
        {
            return groupName switch
            {
                "thu nhập" => 1,
                "kinh doanh" => 2,
                "kinh doanh bancas" => 3,
                _ => 4
            };
        }
    }
}
