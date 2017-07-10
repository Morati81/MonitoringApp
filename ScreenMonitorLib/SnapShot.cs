/*
 * Please leave this Copyright notice in your code if you use it
 * Written by Decebal Mihailescu [http://www.codeproject.com/script/articles/list_articles.asp?userid=634640]
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Reflection;
using System.IO;
using System.Security.Permissions;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Xml;
using System.Security.AccessControl;
using System.Threading;
using WndList = System.Collections.Generic.List<System.IntPtr>;

[assembly: SecurityPermissionAttribute(SecurityAction.RequestMinimum, UnmanagedCode = true)]
[assembly: PermissionSetAttribute(SecurityAction.RequestMinimum, Name = "FullTrust")]

namespace ScreenMonitorLib
{
    public static class Enum<T>
    {
        public static T Parse(string value)
        {
            return (T)Enum.Parse(typeof(T), value);
        }

        public static IList<T> GetValues()
        {
            IList<T> list = new List<T>();
            foreach (object value in Enum.GetValues(typeof(T)))
            {
                list.Add((T)value);
            }
            return list;
        }
    }


    [DebuggerDisplay("lenght = {_tblSnapShots.Rows.Count}")]
    public class SnapShot
    {
        
        private static int _winLong;
        string _folder;
        string _fileName;
        public readonly bool _isService;
        SnapShotDS.WndSettingsDataTable _tblWndSettings;

        public SnapShotDS.WndSettingsDataTable WndSettings
        {
            get { return _tblWndSettings; }
        }
        readonly string _xmlSettings;

        
        public SnapShot(string folder,string fileName)
        {
            
            _isService = true; 
            ScreenMonitorLib.SnapShotDS ds = new ScreenMonitorLib.SnapShotDS();
            _tblWndSettings = ds.WndSettings;
            _folder = folder;
            _fileName = fileName;
            
            _xmlSettings = System.IO.Path.Combine(_folder, "settings.xml");
            // use attributes
            foreach (DataTable tbl in ds.Tables)
            {
                foreach (DataColumn col in tbl.Columns)
                {
                    col.ColumnMapping = MappingType.Attribute;
                }
            }

            try
            {

                _tblWndSettings.CaseSensitive = false;
                if (File.Exists(_xmlSettings))
                    _tblWndSettings.ReadXml(_xmlSettings);
            }


            catch (Exception ex)
            {
                EventLog.WriteEntry("Screen Monitor", string.Format("Unable to read the settings file:\n{0}\n{1}", _xmlSettings, ex.Message),EventLogEntryType.Error, 1, 1);

            }


        }
        

        private static void ExitSpecialCapturing(IntPtr hWnd)
        {
            //EventLog.WriteEntry("Screen Monitor", "in ExitSpecialCapturing winLong =" + _winLong.ToString(), EventLogEntryType.Information, 1, 1);
            if (!Win32API.ShowWindow(hWnd, Win32API.WindowShowStyle.Minimize))
            {
                Win32Exception ex = new Win32Exception(Marshal.GetLastWin32Error());
                throw new ApplicationException("ShowWindow failed: " + ex.Message, ex);
            }
            int res = Win32API.SetWindowLong(hWnd, Win32API.GWL_EXSTYLE, _winLong);
            if (res != (_winLong | Win32API.WS_EX_LAYERED))
            {
                Win32Exception ex = new Win32Exception(Marshal.GetLastWin32Error());
                throw new ApplicationException("SetWindowLong failed: " + ex.Message, ex);
            }
            Win32API.SetMinimizeMaximizeAnimation(true);
        }
        private static void EnterSpecialCapturing(IntPtr hWnd)
        {

            try
            {
                Win32API.SetMinimizeMaximizeAnimation(false);

                _winLong = Win32API.GetWindowLong(hWnd, Win32API.GWL_EXSTYLE);
                if (_winLong == 0)
                {
                    Win32Exception ex = new Win32Exception(Marshal.GetLastWin32Error());
                    throw new ApplicationException("GetWindowLong failed: " + ex.Message, ex);
                }
                int res = Win32API.SetWindowLong(hWnd, Win32API.GWL_EXSTYLE, _winLong | Win32API.WS_EX_LAYERED);
                if (res != _winLong)
                {
                    Win32Exception ex = new Win32Exception(Marshal.GetLastWin32Error());
                    throw new ApplicationException("SetWindowLong failed: " + ex.Message, ex);
                }

                if (!Win32API.SetLayeredWindowAttributes(hWnd, 0, 1, Win32API.LWA_ALPHA))
                {
                    Win32Exception ex = new Win32Exception(Marshal.GetLastWin32Error());
                    throw new ApplicationException("SetLayeredWindowAttributes failed: " + ex.Message, ex);
                }
                Win32API.ShowWindow(hWnd, Win32API.WindowShowStyle.Restore);
                Win32API.SendMessage(hWnd, Win32API.WM_PAINT, 0, 0);
            }
            catch (Exception ex)
            {

                EventLog.WriteEntry("Screen Monitor", ex.Message, EventLogEntryType.Error, 1, 1);
            }
        }

        /// ////////////////////////////////////////////////////////////////
        /// 
        private static bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam)
        {


            if (hWnd == IntPtr.Zero || !Win32API.IsWindow(hWnd)) //|| !Win32API.IsWindowVisible(hWnd))
            {
                throw new ApplicationException("bad window");
            }

            if (!Win32API.IsWindowVisible(hWnd))
                return true;
            GCHandle gch = GCHandle.FromIntPtr(lParam);
            WndList list = gch.Target as WndList;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(hWnd);
            return true;
        }

        bool WndFilter(IntPtr hWnd)
        {
            StringBuilder sb = null;
            try
            {
                sb = new StringBuilder(256);
                Win32API.GetClassName(hWnd, sb, 255);
                string match = sb.ToString();
                sb.Length = 0;
                SnapShotDS.WndSettingsRow rw = _tblWndSettings.FindByClassName(match);
                if (rw == null)
                    return false;
                if (rw.ForegroundOnly && !IsForegroundWnd(hWnd))
                    return false;
                if (!rw.GetIconicWnd && Win32API.IsIconic(hWnd))
                    return false;
                return true;
            }
            finally
            {
                sb.Length = 0;
            }
        }

        public WndList GetDesktopWindows(IntPtr hDesktop)
        {
            WndList lst = new WndList();
            GCHandle listHandle = default(GCHandle);
            listHandle = GCHandle.Alloc(lst);

            Win32API.EnumDelegate enumfunc = new Win32API.EnumDelegate(EnumWindowsProc);
            //IntPtr hDesktop = IntPtr.Zero; // current desktop
            try
            {
                bool success = Win32API.EnumDesktopWindows(hDesktop, enumfunc, GCHandle.ToIntPtr(listHandle));

                if (success)
                {

                    return lst.FindAll(new Predicate<IntPtr>(WndFilter));

                }
                else
                {
                    // Get the last Win32 error code
                    int errorCode = Marshal.GetLastWin32Error();
                    Win32Exception ex = new Win32Exception(errorCode);
                    string errorMessage = String.Format("EnumDesktopWindows failed with code {0}.\n {1}", errorCode, ex.Message);
                    throw new ApplicationException(errorMessage, ex);
                }
            }
            finally
            {

                if (listHandle != default(GCHandle) && listHandle.IsAllocated)
                    listHandle.Free();
            }
        }
       

        public void SaveSettings()
        {
            _tblWndSettings.AcceptChanges();
            SnapShotDS.WndSettingsRow rw = _tblWndSettings.FindByClassName(string.Empty);
            if (rw != null)
                _tblWndSettings.Rows.Remove(rw);
            _tblWndSettings.AcceptChanges();
            _tblWndSettings.WriteXml(_xmlSettings);
        }

        public void SaveSnapShots(WndList lst)
        {
            try
            {
                
                lst.ForEach(delegate (IntPtr hWnd) { SaveSnapShot(hWnd); });
                
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("Screen Monitor", string.Format("dataset exception in SaveSnapShots:{0} when isservice = {1}", ex.Message, _isService),
                     EventLogEntryType.Error, 1, 1);
            }
            finally
            {
              
            }
        }
        /// ////////////////////////////////////////////////////////////////
        public void SaveAllSnapShots(WndList lst)
        {
            if (!_isService)
                return;

            WndList iconiclst = lst.FindAll(delegate(IntPtr hWnd) { return Win32API.IsIconic(hWnd); });
            lst.RemoveAll(delegate(IntPtr hWnd) { return Win32API.IsIconic(hWnd); });
            //will create history file first as service
            SaveSnapShots(lst);
            lst.Clear();
          

        }

        

        static bool IsForegroundWnd(IntPtr hWnd)
        {
            return Win32API.GetForegroundWindow() == hWnd;
        }

        public bool SaveSnapShot(IntPtr hWnd)
        {
            Bitmap bitmap = null;
            int procId = 0;
            int threadId = 0;
            IntPtr hOriginalFGWnd = IntPtr.Zero;
            IntPtr hOriginalFocusWnd = IntPtr.Zero;
            try
            {
                hOriginalFGWnd = Win32API.GetForegroundWindow();
                threadId = Win32API.GetWindowThreadProcessId(hWnd, out procId);
                if (procId > 0)
                {
                    if (!Win32API.AttachThreadInput(Win32API.GetCurrentThreadId(), threadId, true))
                        EventLog.WriteEntry("Screen Monitor", string.Format("failed to attach{0} to {1} is service = {2}", Win32API.GetCurrentThreadId(), threadId, _isService),
                            EventLogEntryType.Error, 1, 1);
                    else
                    {
                        hOriginalFocusWnd = Win32API.GetFocus();
                    }

                }

                if (hWnd == IntPtr.Zero || !Win32API.IsWindow(hWnd) || !Win32API.IsWindowVisible(hWnd))
                {
                    EventLog.WriteEntry("Screen Monitor", "unusable window ", EventLogEntryType.Error, 1, 1);
                    return false;
                }

                bool isIconic = Win32API.IsIconic(hWnd);

                StringBuilder sb = new StringBuilder(256);
                Win32API.GetClassName(hWnd, sb, 255);
                string match = sb.ToString();
                SnapShotDS.WndSettingsRow rowSettings = _tblWndSettings.FindByClassName(match);

                bitmap = MakeSnapshot(hWnd, rowSettings.ClientWindow, Win32API.WindowShowStyle.Restore);


                if (bitmap == null)
                {
                    return false;
                }

                PersistCapture(hWnd, bitmap, isIconic, rowSettings);
                return true;
            }
            catch (Exception ex)
            {

                EventLog.WriteEntry("Screen Monitor", "SaveSnapShot failed: " + ex.Message, EventLogEntryType.Error, 1, 1);
                return false;
            }
            finally
            {
                if (bitmap != null)
                    bitmap.Dispose();
                if (hOriginalFGWnd != Win32API.GetForegroundWindow())
                    Win32API.SetForegroundWindow(hOriginalFGWnd);
                if (procId > 0)
                {
                    if (hOriginalFocusWnd != IntPtr.Zero)
                    {
                        if (IntPtr.Zero == Win32API.SetFocus(hOriginalFocusWnd))
                        {
                            Win32Exception ex = new Win32Exception(Marshal.GetLastWin32Error());
                            EventLog.WriteEntry("Screen Monitor", string.Format("SetFocus for {0} failed with code {1}: {2}", hOriginalFocusWnd, ex.ErrorCode, ex.Message), EventLogEntryType.Error, 1, 1);

                        }
                        if (!Win32API.AttachThreadInput(Win32API.GetCurrentThreadId(), threadId, false))
                            EventLog.WriteEntry("Screen Monitor", string.Format("failed to attach {0} from {1} is service = {2}", Win32API.GetCurrentThreadId(), threadId, _isService),
        EventLogEntryType.Error, 1, 1);
                    }


                }
            }


        }

        private void PersistCapture(IntPtr hWnd, Bitmap bitmap, bool isIconic, SnapShotDS.WndSettingsRow rowSettings)
        {
            using (MemoryStream ms = new MemoryStream())
            {

                bitmap.Save(ms, ImageFormat.Jpeg);
                MD5 md5 = new MD5CryptoServiceProvider();
                md5.Initialize();
                ms.Position = 0;
                byte[] result = md5.ComputeHash(ms);
                Guid guid = new Guid(result);
                //int len = _tblSnapShots.Select(string.Format("{0} = '{1}'", _tblSnapShots.FileNameColumn.ColumnName, guid.ToString())).Length;
                
                string path = System.IO.Path.Combine(_folder, _fileName);
                //if (len == 0 || !File.Exists(path))
                    if (!File.Exists(path))
                    {

                    using (FileStream fs = File.OpenWrite(path))
                    {
                        ms.WriteTo(fs);
                    }
                }

               

            }
        }

        System.Drawing.Bitmap MakeSnapshot(IntPtr hWnd, bool isClient, Win32API.WindowShowStyle nCmdShow)
        {
            //paint control onto graphics using provided options  
            IntPtr hDC = IntPtr.Zero;
            IntPtr hdcTo = IntPtr.Zero;
            IntPtr hBitmap = IntPtr.Zero;
            bool bIsiconic = false;
            if (hWnd == IntPtr.Zero || !Win32API.IsWindow(hWnd) || !Win32API.IsWindowVisible(hWnd))
                return null;

            try
            {


                if (Win32API.IsIconic(hWnd))
                {
                    if (_isService)
                        Win32API.ShowWindow(hWnd, nCmdShow);
                    else
                    {
                        bIsiconic = true;
                        EnterSpecialCapturing(hWnd);
                    }
                }

            }
            catch (Exception ex)
            {

                EventLog.WriteEntry("Screen Monitor", "EnterSpecialCapturing failed " + ex.Message, EventLogEntryType.Warning, 1, 1);
            }


            System.Drawing.Bitmap image = null;
            RECT appRect = new RECT();
            System.Drawing.Graphics graphics = null;
            try
            {
                Win32API.GetWindowRect(hWnd, out appRect);

                image = new System.Drawing.Bitmap(appRect.Width, appRect.Height);

                graphics = System.Drawing.Graphics.FromImage(image);
                hDC = graphics.GetHdc();



                Win32API.PrintWindow(hWnd, hDC, 0);//Win32API.PW_CLIENTONLY);
                if (!isClient)
                    return image;
                RECT clientRect;
                bool res = Win32API.GetClientRect(hWnd, out clientRect);
                Point lt = new Point(clientRect.Left, clientRect.Top);
                Win32API.ClientToScreen(hWnd, ref lt);

                hdcTo = Win32API.CreateCompatibleDC(hDC);
                hBitmap = Win32API.CreateCompatibleBitmap(hDC, clientRect.Width, clientRect.Height);

                //  validate...
                if (hBitmap != IntPtr.Zero)
                {
                    // copy...
                    int x = lt.X - appRect.Left;
                    int y = lt.Y - appRect.Top;
                    IntPtr hLocalBitmap = Win32API.SelectObject(hdcTo, hBitmap);

                    Win32API.BitBlt(hdcTo, 0, 0, clientRect.Width, clientRect.Height, hDC, x, y, Win32API.SRCCOPY);
                    //  create bitmap for window image...
                    image.Dispose();
                    image = System.Drawing.Image.FromHbitmap(hBitmap);
                }

            }

            finally
            {
                if (hBitmap != IntPtr.Zero)
                    Win32API.DeleteObject(hBitmap);
                if (hdcTo != IntPtr.Zero)
                    Win32API.DeleteDC(hdcTo);
                graphics.ReleaseHdc(hDC);
                if (bIsiconic)
                    ExitSpecialCapturing(hWnd);
            }

            return image;

        }
    }
}
