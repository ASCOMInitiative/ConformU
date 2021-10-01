using System.Collections.Generic;

namespace ConformU
{
    /// <summary>
    /// Blazorise Tree control item
    /// </summary>
    public class BlazoriseTreeItem
    {
        public string Text { get; set; }
        public IEnumerable<BlazoriseTreeItem> Children { get; set; }
    }
}
