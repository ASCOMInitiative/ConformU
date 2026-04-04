using System.Collections.Generic;

namespace ConformU
{
    /// <summary>
    /// Tree control item
    /// </summary>
    public class TreeItem
    {
        public string Text { get; set; }
        public IEnumerable<TreeItem> Children { get; set; }
    }
}
