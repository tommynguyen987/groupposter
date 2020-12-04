using System;

namespace NTPEmailMarketing
{    
    public class HttpSession
    {
        //This static parameter is set true when we do in fresh price, we need the time out for one request
        public static bool _bIsNeedTimeOut = false;
        //This is the IP which use in downloading
        public static string sIP = "-1";
        private string sContent = "";
        private System.Net.CookieCollection _cookies = new System.Net.CookieCollection();
        private System.Net.WebHeaderCollection _headers = new System.Net.WebHeaderCollection();
        public string sReferer = "";
        public string sAccept = "";
        public string sContentType = "";
        public int _iTimeOut = 30 * 1000; //Time out is 30s
        private int iDefaultTimeout = 60 * 1000; // Default time out is 60s

        public HttpSession()
        {
            //Read IP from config file
            if (!String.IsNullOrEmpty(sIP))
                if (sIP == "-1")
                    sIP = "";// System.Configuration.ConfigurationManager.AppSettings["IP"];
        }
        public System.Net.CookieCollection Cookies
        {
            get { return _cookies; }
            set { _cookies = value; }
        }
        public System.Net.WebHeaderCollection Headers
        {
            get { return _headers; }
            set { _headers = value; }
        }
        public string GetMethodDownload(string sUrl, bool isKeepAlive, bool isAutoRedirect, bool isUseHeader, bool isUseCookie)
        {
            try
            {
                sContent = "";
                System.Net.HttpWebRequest httpRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(sUrl);
                //Begin check to use default IP or use IP from config file
                if ((String.IsNullOrEmpty(sIP) == false) && (sIP != "-1"))
                    httpRequest.ServicePoint.BindIPEndPointDelegate = new System.Net.BindIPEndPoint(Bind);
                //End check
                httpRequest.Method = "GET";
                if (!sContentType.Equals(""))
                    httpRequest.ContentType = sContentType;
                else
                    httpRequest.ContentType = "application/x-www-form-urlencoded";
                if (!sReferer.Equals(""))
                    httpRequest.Referer = sReferer;
                httpRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows XP)";
                httpRequest.AllowAutoRedirect = isAutoRedirect;
                httpRequest.KeepAlive = isKeepAlive;
                httpRequest.ServicePoint.ConnectionLimit = 500;
                httpRequest.Credentials = System.Net.CredentialCache.DefaultCredentials;
                httpRequest.CookieContainer = new System.Net.CookieContainer();
                if (!string.IsNullOrEmpty(sAccept))
                    httpRequest.Accept = sAccept;

                if (_bIsNeedTimeOut) httpRequest.Timeout = _iTimeOut;
                else httpRequest.Timeout = iDefaultTimeout;
                if (_cookies != null && isUseCookie == true)
                {
                    httpRequest.CookieContainer.Add(_cookies);
                }
                if (_headers != null & isUseHeader == true)
                {
                    string key;
                    string keyvalue;
                    for (int i = 0; i < _headers.Count; i++)
                    {

                        key = _headers.Keys[i];
                        keyvalue = _headers[i];
                        httpRequest.Headers.Add(key, keyvalue);
                    }
                }

                //Add gzip 
                if (httpRequest.Headers == null)
                {
                    httpRequest.Headers = new System.Net.WebHeaderCollection();
                }
                httpRequest.Headers.Add("Accept-Encoding", "gzip,deflate");

                System.Net.HttpWebResponse httpResponse = (System.Net.HttpWebResponse)httpRequest.GetResponse();
                //string sCharacterSet = "UTF-8";
                //if (httpResponse.CharacterSet != "") sCharacterSet = httpResponse.CharacterSet;
                //Encoding encoding = Encoding.GetEncoding(sCharacterSet);
                System.IO.Stream answer = httpResponse.GetResponseStream();

                //Unzip
                if (httpResponse.ContentEncoding.ToLower().Contains("gzip"))
                {
                    answer = new System.IO.Compression.GZipStream(answer, System.IO.Compression.CompressionMode.Decompress);
                }
                else if (httpResponse.ContentEncoding.ToLower().Contains("deflate"))
                {
                    answer = new System.IO.Compression.DeflateStream(answer, System.IO.Compression.CompressionMode.Decompress);
                }
                //
                System.IO.StreamReader _answer = new System.IO.StreamReader(answer);
                _headers = httpResponse.Headers;
                foreach (System.Net.Cookie cook in httpResponse.Cookies)
                    _cookies.Add(cook);
                //_cookies = httpResponse.Cookies;
                sContent = _answer.ReadToEnd();
                httpResponse.Close();
            }
            catch (Exception e)
            {
                /*T  temp remove Microsoft.Practices.EnterpriseLibrary.Logging */

                //LogEntry en = new LogEntry();
                //en.Message = "ERROR: could not download content from url:" + sUrl + "--" + e.Message;
                //Logger.Write(en);
                System.Diagnostics.Trace.WriteLine("ERROR: could not download content from url:" + sUrl + "--" + e.Message);
                //throw e;
            }
            return sContent;
        }
        public string PostMethodDownload(string sUrl, string sPost, bool isKeepAlive, bool isAutoRedirect, bool isUseHeader, bool isUseCookie)
        {
            try
            {
                sContent = "";
                byte[] buffer = System.Text.Encoding.ASCII.GetBytes(sPost);

                System.Net.HttpWebRequest httpRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(sUrl);
                //Begin check to use default IP or use IP from config file
                if ((String.IsNullOrEmpty(sIP) == false) && (sIP != "-1"))
                    httpRequest.ServicePoint.BindIPEndPointDelegate = new System.Net.BindIPEndPoint(Bind);
                //End check

                httpRequest.Method = "POST";
                httpRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows XP)";
                if (!sContentType.Equals(""))
                    httpRequest.ContentType = sContentType;
                else
                    httpRequest.ContentType = "application/x-www-form-urlencoded";
                if (!sReferer.Equals(""))
                    httpRequest.Referer = sReferer;
                if (_bIsNeedTimeOut) httpRequest.Timeout = _iTimeOut;
                else httpRequest.Timeout = iDefaultTimeout;
                httpRequest.AllowAutoRedirect = isAutoRedirect;
                httpRequest.KeepAlive = isKeepAlive;
                httpRequest.ServicePoint.ConnectionLimit = 500;
                httpRequest.ContentLength = buffer.Length;
                httpRequest.CookieContainer = new System.Net.CookieContainer();


                if (_cookies != null && isUseCookie == true)
                {
                    httpRequest.CookieContainer.Add(_cookies);
                }
                if (_headers != null & isUseHeader == true)
                {
                    string key;
                    string keyvalue;
                    for (int i = 0; i < _headers.Count; i++)
                    {

                        key = _headers.Keys[i];
                        keyvalue = _headers[i];
                        httpRequest.Headers.Add(key, keyvalue);
                    }
                }
                //Add gzip 
                if (httpRequest.Headers == null)
                {
                    httpRequest.Headers = new System.Net.WebHeaderCollection();
                }
                httpRequest.Headers.Add("Accept-Encoding", "gzip,deflate");
                //
                System.IO.Stream postData = httpRequest.GetRequestStream();
                postData.Write(buffer, 0, buffer.Length);
                postData.Close();
                System.Net.HttpWebResponse httpResponse = (System.Net.HttpWebResponse)httpRequest.GetResponse();
                //string sCharacterSet = "UTF-8";
                //if (httpResponse.CharacterSet != "") sCharacterSet = httpResponse.CharacterSet;
                //Encoding encoding = Encoding.GetEncoding(sCharacterSet);
                System.IO.Stream answer = httpResponse.GetResponseStream();

                //Unzip
                if (httpResponse.ContentEncoding.ToLower().Contains("gzip"))
                {
                    answer = new System.IO.Compression.GZipStream(answer, System.IO.Compression.CompressionMode.Decompress);
                }
                else if (httpResponse.ContentEncoding.ToLower().Contains("deflate"))
                {
                    answer = new System.IO.Compression.DeflateStream(answer, System.IO.Compression.CompressionMode.Decompress);
                }
                foreach (System.Net.Cookie cook in httpResponse.Cookies)
                    _cookies.Add(cook);
                //_cookies = httpResponse.Cookies;
                _headers = httpResponse.Headers;
                // StreamReader _answer = new StreamReader(answer, encoding);
                System.IO.StreamReader _answer = new System.IO.StreamReader(answer);
                sContent = _answer.ReadToEnd();
                httpResponse.Close();
            }            
            catch (Exception e)
            {
                /*T temp remove Microsoft.Practices.EnterpriseLibrary.Logging*/
                //LogEntry en = new LogEntry();
                //en.Message = "ERROR: could not download content from url:" + sUrl + "--" + e.Message;
                //Logger.Write(en);
               sContent="ERROR: could not download content from url:" + sUrl + "--" + e.Message;
            }            
            return sContent;
        }
        System.Net.IPEndPoint Bind(System.Net.ServicePoint servicePoint, System.Net.IPEndPoint remoteEndPoint, int retryCount)
        {
            System.Net.IPAddress address = System.Net.IPAddress.Parse(sIP); //Replace with any of your machine's external IPs
            //The second parameter is 0: meaning that the client port will be assigned by windows
            System.Net.IPEndPoint ie = new System.Net.IPEndPoint(address, 0);
            return ie;
        }        
    }
}