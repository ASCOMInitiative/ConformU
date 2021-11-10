using Microsoft.JSInterop;
using System;

namespace ConformU
{
    /// <summary>
    /// React to browser resize events sent from the JavaScript listener in _Host.cshtml
    /// </summary>
    public class BrowserResizeService
    {
        private IJSRuntime JS = null;

        public event EventHandler<BrowserWindowSize> OnResize;

        public async void Init(IJSRuntime js)
        {
            // Enforce single invocation            
            if (JS == null)
            {
                this.JS = js;
                await JS.InvokeAsync<string>("resizeListener", DotNetObjectReference.Create(this));
            }
        }

        [JSInvokable]
        public void SetBrowserDimensions(int jsBrowserWidth, int jsBrowserHeight)
        {
            BrowserWindowSize browserWindowSize = new();
            browserWindowSize.Width = jsBrowserWidth;
            browserWindowSize.Height = jsBrowserHeight;
            // For simplicity, we're just using the new width
            this.OnResize?.Invoke(this, browserWindowSize);
        }
    }
}