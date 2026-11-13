using System;
using System.Runtime.InteropServices;

[assembly: ComVisible(false)]
[assembly: System.Reflection.AssemblyVersion("1.1.0.0")]

namespace EZOffice
{
    // Office calls these COM interfaces by GUID. Keeping their declarations here
    // lets the add-in build without Visual Studio or the Office interop assemblies.
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application,
            int connectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object addIn,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(int removeMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    [ComVisible(true), Guid("DDB76EDA-0745-45D4-B0DB-331CA329237A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonCallbacks
    {
        [DispId(1)]
        void RibbonLoaded([MarshalAs(UnmanagedType.IDispatch)] object ribbon);

        [DispId(2)]
        [return: MarshalAs(UnmanagedType.VariantBool)]
        bool GetScrollEnabled([MarshalAs(UnmanagedType.IDispatch)] object control);

        [DispId(3)]
        void SetScrollEnabled([MarshalAs(UnmanagedType.IDispatch)] object control,
            [MarshalAs(UnmanagedType.VariantBool)] bool enabled);

        [DispId(4)]
        [return: MarshalAs(UnmanagedType.VariantBool)]
        bool GetAutoSaveEnabled([MarshalAs(UnmanagedType.IDispatch)] object control);

        [DispId(5)]
        void SetAutoSaveEnabled([MarshalAs(UnmanagedType.IDispatch)] object control,
            [MarshalAs(UnmanagedType.VariantBool)] bool enabled);

        [DispId(6)]
        string GetInterval([MarshalAs(UnmanagedType.IDispatch)] object control);

        [DispId(7)]
        void SetInterval([MarshalAs(UnmanagedType.IDispatch)] object control, string text);

        [DispId(8)]
        string GetSaveStatus([MarshalAs(UnmanagedType.IDispatch)] object control);

        [DispId(9)]
        string GetScrollLabel([MarshalAs(UnmanagedType.IDispatch)] object control);
    }
}
