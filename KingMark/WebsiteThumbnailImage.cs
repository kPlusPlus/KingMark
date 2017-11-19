using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace MohammadDayyan
{

    /// <summary>
    /// Generates websites thumbnail image
    /// Learn more in http://www.beansoftware.com/ASP.NET-Tutorials/Get-Web-Site-Thumbnail-Image.aspx
    /// </summary>
    class WebsiteThumbnailImage
    {
        string m_Url = "";
        int m_BrowserWidth, m_BrowserHeight, m_ThumbnailWidth, m_ThumbnailHeight;
        Bitmap m_Bitmap;                

        public WebsiteThumbnailImage(string Url, int BrowserWidth, int BrowserHeight, int ThumbnailWidth, int ThumbnailHeight)
        {
            this.m_Url = Url;
            this.m_BrowserWidth = BrowserWidth;
            this.m_BrowserHeight = BrowserHeight;
            this.m_ThumbnailWidth = ThumbnailWidth;
            this.m_ThumbnailHeight = ThumbnailHeight;
        }

        public Bitmap GenerateWebSiteThumbnailImage()
        {
            Thread m_thread = new Thread(new ThreadStart(_GenerateWebSiteThumbnailImage));            
            m_thread.IsBackground = true;
            m_thread.SetApartmentState(ApartmentState.STA);
            m_thread.Start();            
            m_thread.Join();
            return m_Bitmap;
        }

        private void _GenerateWebSiteThumbnailImage()
        {
            //if the thread doesn't finish its work on maxTime, thread will be abort
            short maxTime = 2 * 60;//2 minutes

            WebBrowser m_WebBrowser = new WebBrowser();
            m_WebBrowser.ScrollBarsEnabled = false;
            m_WebBrowser.Navigate(m_Url);
            m_WebBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(WebBrowser_DocumentCompleted);
            while (m_WebBrowser.ReadyState != WebBrowserReadyState.Complete && maxTime > 0)
            {
                Thread.Sleep(1 * 1000);
                maxTime--;
                Application.DoEvents();
            }
            m_WebBrowser.Dispose();
        }

        private void WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser m_WebBrowser = (WebBrowser)sender;
            m_WebBrowser.ClientSize = new Size(this.m_BrowserWidth, this.m_BrowserHeight);
            m_WebBrowser.ScrollBarsEnabled = false;
            m_Bitmap = new Bitmap(m_WebBrowser.Bounds.Width, m_WebBrowser.Bounds.Height);
            m_WebBrowser.BringToFront();
            m_WebBrowser.DrawToBitmap(m_Bitmap, m_WebBrowser.Bounds);
            m_Bitmap = (Bitmap)m_Bitmap.GetThumbnailImage(m_ThumbnailWidth, m_ThumbnailHeight, null, IntPtr.Zero);
        }
    }
}
