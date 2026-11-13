using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.CSharp.RuntimeBinder;

namespace EZOffice
{
    [ComVisible(true), Guid("717C59F4-59F8-4A69-87F0-02BFA563FE98"), ProgId("EZOffice.AddIn")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IRibbonCallbacks))]
    public sealed class AddIn : IDTExtensibility2, IRibbonExtensibility, IRibbonCallbacks
    {
        private const int DefaultMinutes = 5;
        private OfficeHost host;
        private string settingsKey;
        private dynamic ribbon;
        private HorizontalScroll scroll;
        private Timer saveTimer;
        private bool autoSaveEnabled;
        private bool scrollUnavailable;
        private bool saving;
        private int intervalMinutes = DefaultMinutes;
        private string saveStatus = "自动保存已关闭";

        public void OnConnection(object application, int connectMode, object addIn, ref Array custom)
        {
            host = new OfficeHost(application);
            settingsKey = host.Name == "Word" ? @"Software\EZOffice" : @"Software\EZOffice\" + host.Name;
            saveTimer = new Timer();
            saveTimer.Tick += SaveTick;
            autoSaveEnabled = ReadSetting("AutoSave", 0) != 0;
            int storedInterval = ReadSetting("IntervalMinutes", DefaultMinutes);
            intervalMinutes = IsValidInterval(storedInterval) ? storedInterval : DefaultMinutes;
            SetScrollState(ReadSetting("HorizontalScroll", 1) != 0);
            RestartTimer();
        }

        public void OnDisconnection(int removeMode, ref Array custom) { Shutdown(); }
        public void OnBeginShutdown(ref Array custom) { Shutdown(); }
        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }

        public string GetCustomUI(string ribbonId)
        {
            using (var reader = new StreamReader(typeof(AddIn).Assembly.GetManifestResourceStream("EZOffice.Ribbon.xml")))
                return reader.ReadToEnd();
        }

        public void RibbonLoaded(object ribbonUi) { ribbon = ribbonUi; }
        public bool GetScrollEnabled(object control) { return scroll != null; }
        public string GetScrollLabel(object control)
        {
            return scrollUnavailable ? "横向滚动（启用失败）" : "横向滚动";
        }

        public void SetScrollEnabled(object control, bool enabled)
        {
            SetScrollState(enabled);
            WriteSetting("HorizontalScroll", scroll != null ? 1 : 0);
            Invalidate("EZOfficeScroll");
        }

        public bool GetAutoSaveEnabled(object control) { return autoSaveEnabled; }

        public void SetAutoSaveEnabled(object control, bool enabled)
        {
            autoSaveEnabled = enabled;
            WriteSetting("AutoSave", enabled ? 1 : 0);
            RestartTimer();
            Invalidate("EZOfficeAutoSave");
        }

        public string GetInterval(object control) { return intervalMinutes.ToString(CultureInfo.InvariantCulture); }

        public void SetInterval(object control, string text)
        {
            int minutes;
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out minutes) && IsValidInterval(minutes))
            {
                intervalMinutes = minutes;
                WriteSetting("IntervalMinutes", minutes);
                RestartTimer();
            }
            else
            {
                SetSaveStatus("间隔须为 1–120 分钟");
            }
            Invalidate("EZOfficeInterval");
        }

        public string GetSaveStatus(object control) { return saveStatus; }

        private void SetScrollState(bool enabled)
        {
            scrollUnavailable = false;
            if (!enabled && scroll != null)
            {
                scroll.Dispose();
                scroll = null;
            }
            if (enabled && scroll == null)
            {
                try { scroll = new HorizontalScroll(host.FrameClass, host.EditorClass, Scroll); }
                catch (System.ComponentModel.Win32Exception error)
                {
                    Trace.WriteLine(error);
                    scrollUnavailable = true;
                }
            }
        }

        private bool Scroll(IntPtr frame, int lines, double editorX, double editorY)
        {
            if (host == null || scroll == null)
                return false;

            try { return host.TryScroll(frame, lines, editorX, editorY); }
            catch (COMException error) { Trace.WriteLine(error); }
            catch (RuntimeBinderException error) { Trace.WriteLine(error); }
            return false;
        }

        private static bool IsValidInterval(int minutes) { return minutes >= 1 && minutes <= 120; }

        private void RestartTimer()
        {
            saveTimer.Stop();
            saveTimer.Interval = intervalMinutes * 60000;
            if (autoSaveEnabled && !saving)
                saveTimer.Start();
            SetSaveStatus(autoSaveEnabled ? "等待定时保存" : "自动保存已关闭");
        }

        private void SaveTick(object sender, EventArgs args)
        {
            // Office can pump messages while saving. Keep a second timer tick out
            // until this pass completes, then start the next interval.
            saving = true;
            saveTimer.Stop();
            try
            {
                if (HorizontalScroll.IsInModalDialog(host.FrameClass))
                    return;

                string result = LocalAutoSave.SaveChangedFiles(host.GetFilesReadyForSaving(),
                    () => autoSaveEnabled && saveTimer != null);
                if (result != null && autoSaveEnabled && saveTimer != null)
                    SetSaveStatus(result);
            }
            catch (COMException error)
            {
                Trace.WriteLine(error);
                if (autoSaveEnabled && saveTimer != null)
                    SetSaveStatus("应用忙碌，下次重试");
            }
            finally
            {
                saving = false;
                if (autoSaveEnabled && saveTimer != null)
                    saveTimer.Start();
            }
        }

        private void SetSaveStatus(string text)
        {
            saveStatus = text;
            Invalidate("EZOfficeSaveStatus");
        }

        private void Invalidate(string control)
        {
            if (ribbon == null)
                return;
            try { ribbon.InvalidateControl(control); }
            catch (COMException error) { Trace.WriteLine(error); }
        }

        private void Shutdown()
        {
            if (saveTimer != null)
            {
                saveTimer.Dispose();
                saveTimer = null;
            }
            SetScrollState(false);
            ribbon = null;
            host = null;
        }

        private int ReadSetting(string name, int fallback)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(settingsKey))
                {
                    if (key != null && key.GetValue(name) is int)
                        return (int)key.GetValue(name);
                }

                string legacyKey = settingsKey.Replace(@"Software\EZOffice", @"Software\EZword");
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(legacyKey))
                {
                    int? value = key == null ? null : key.GetValue(name) as int?;
                    if (value.HasValue)
                        WriteSetting(name, value.Value);
                    return value ?? fallback;
                }
            }
            catch (SystemException error)
            {
                Trace.WriteLine(error);
                return fallback;
            }
        }

        private void WriteSetting(string name, int value)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(settingsKey))
                    key.SetValue(name, value, RegistryValueKind.DWord);
            }
            catch (SystemException error) { Trace.WriteLine(error); }
        }
    }
}
