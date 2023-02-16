using System.Diagnostics;
using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.UI.WindowsAndMessaging;
using Windows.Win32.UI.Shell;
using Windows.Win32.UI.Shell.Common;
using Windows.Win32.System.Registry;
using Windows.Win32.UI.Input.KeyboardAndMouse;
using Microsoft.Win32.SafeHandles;

namespace WinExplorer
{
    using static Windows.Win32.PInvoke;

    public class WinExplorerLikeContextMenu : IDisposable
    {
        const int CMIC_MASK_ICON           = 0x00000010;
        const int CMIC_MASK_HOTKEY         = 0x00000020;
        const int CMIC_MASK_NOASYNC        = 0x00000100;
        const int CMIC_MASK_FLAG_NO_UI     = 0x00000400;
        const int CMIC_MASK_UNICODE        = 0x00004000;
        const int CMIC_MASK_NO_CONSOLE     = 0x00008000;
        const int CMIC_MASK_ASYNCOK        = 0x00100000;
        const int CMIC_MASK_NOZONECHECKS   = 0x00800000;
        const int CMIC_MASK_FLAG_LOG_USAGE = 0x04000000;

        static string _className = "ExplorerShellTestWinFormsApp.MenuHostWindow";

        static WinExplorerLikeContextMenu? constructingMenuHostWindow;

        static object windowOperationGate = new object();

        static IShellFolder s_desktopFolder;

        HWND _hWnd;

        IContextMenu2? _workingContextMenu2;
        IContextMenu3? _workingContextMenu3;

        static WinExplorerLikeContextMenu()
        {
            var wndClass = new WNDCLASSW();

            unsafe
            {
                fixed (char* className = _className)
                {
                    wndClass.lpfnWndProc = InitWndProc;
                    wndClass.lpszClassName = className;

                    RegisterClass(wndClass);
                }
            }

            SHGetDesktopFolder(out s_desktopFolder).ThrowOnFailure();
        }

        public unsafe WinExplorerLikeContextMenu()
        {
            var moduleHandle = GetModuleHandle(default(PCWSTR));

            fixed (char* className = _className)
            {
                lock (windowOperationGate)
                {
                    Debug.Assert(constructingMenuHostWindow is null);
                    try
                    {
                        constructingMenuHostWindow = this;

                        var hWnd = CreateWindowEx(
                            dwExStyle: 0,
                            lpClassName: className,
                            lpWindowName: null,
                            dwStyle: 0,
                            X: 0,
                            Y: 0,
                            nWidth: 0,
                            nHeight: 0,
                            hWndParent: HWND.Null,
                            hMenu: default,
                            hInstance: moduleHandle,
                            lpParam: null
                            );

                        if (hWnd.IsNull)
                        {
                            throw new InvalidOperationException();
                        }

                        if (hWnd != _hWnd)
                        {
                            throw new InvalidOperationException();
                        }
                    }
                    finally
                    {
                        constructingMenuHostWindow = null;
                    }
                }
            }
        }

        ~WinExplorerLikeContextMenu() => Dispose(disposing: false);

        public unsafe bool Show(int pointScreenX, int pointScreenY, IReadOnlyCollection<FileSystemInfo> fileSystemInfos)
        {
            if (_hWnd.IsNull)
            {
                throw new ObjectDisposedException(null);
            }

            if (GetWindowThreadProcessId(_hWnd) != GetCurrentThreadId())
            {
                throw new InvalidOperationException();
            }

            if (fileSystemInfos.Count == 0)
            {
                return false;
            }

            var parentDir = GetParentDirectoryInfo(fileSystemInfos.First());

            if (parentDir is null)
            {
                return false;
            }

            if (_workingContextMenu2 is not null || _workingContextMenu3 is not null)
            {
                return false;
            }


            HMENU hMenu = HMENU.Null;
            IContextMenu? contextMenu = null;
            try
            {
                var fileSystemInfosWithParent = new (FileSystemInfo fileSytemInfo, DirectoryInfo parent)[fileSystemInfos.Count];

                int i = 0;
                foreach (var fileSytemInfo in fileSystemInfos)
                {
                    var parent = GetParentDirectoryInfo(fileSytemInfo);

                    if (parent is null)
                    {
                        return false;
                    }

                    fileSystemInfosWithParent[i++] = (fileSytemInfo, parent);
                }

                bool needHandleOpenCmd = false;
                bool ignoreEnableExtendedVerbs = false;

                if (fileSystemInfosWithParent.Skip(1).All(f => f.parent.FullName == fileSystemInfosWithParent[0].parent.FullName))
                {
                    contextMenu = CreateSingleDirectoryFileSystemItemsContextMenu(parentDir, fileSystemInfos);
                }
                else
                {
                    contextMenu = CreateMultipleDirectoryFileSystemItemsContextMenu(fileSystemInfos, out needHandleOpenCmd, out ignoreEnableExtendedVerbs);
                }

                _workingContextMenu2 = contextMenu as IContextMenu2;
                _workingContextMenu3 = contextMenu as IContextMenu3;

                hMenu = CreatePopupMenu();
                const int CMD_FIRST = 1;
                const int CMD_LAST = int.MaxValue;

                bool enableExtendedVerbs = GetKeyState((int)VIRTUAL_KEY.VK_SHIFT) < 0;

                contextMenu
                    .QueryContextMenu(hMenu, 0, CMD_FIRST, CMD_LAST, CMF_EXPLORE | CMF_NORMAL | (!ignoreEnableExtendedVerbs && enableExtendedVerbs ? CMF_EXTENDEDVERBS : 0))
                    .ThrowOnFailure();

                var selectedCmd = TrackPopupMenuEx(hMenu, (uint)TRACK_POPUP_MENU_FLAGS.TPM_RETURNCMD, pointScreenX, pointScreenY, _hWnd, null);

                if (selectedCmd.Value > 0)
                {
                    var buffer = new char[1024];

                    fixed (char* bufferPtr = buffer)
                    {
                        contextMenu
                            .GetCommandString((uint)selectedCmd.Value - CMD_FIRST, GCS_VERBW, null, (byte*)bufferPtr, (uint)buffer.Length)
                            .ThrowOnFailure();
                    }

                    var verb = new string(buffer, 0, Array.IndexOf(buffer, '\0'));

                    if (verb == "open" && needHandleOpenCmd)
                    {
                        foreach (var fileSystemInfo in fileSystemInfos)
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                UseShellExecute = true,
                                FileName = fileSystemInfo.FullName,
                            });
                        }
                    }
                    else
                    {
                        fixed (char* directory = parentDir.FullName)
                        {
                            CMINVOKECOMMANDINFOEX invoke = new CMINVOKECOMMANDINFOEX();
                            invoke.cbSize = (uint)Marshal.SizeOf<CMINVOKECOMMANDINFOEX>();
                            invoke.lpVerb = new PCSTR((byte*)new IntPtr(selectedCmd.Value - CMD_FIRST).ToPointer());
                            invoke.lpVerbW = new PCWSTR((char*)new IntPtr(selectedCmd.Value - CMD_FIRST).ToPointer());
                            invoke.lpDirectoryW = directory;
                            invoke.fMask = CMIC_MASK_UNICODE | CMIC_MASK_PTINVOKE
                                | (GetKeyState((int)VIRTUAL_KEY.VK_CONTROL) < 0 ? CMIC_MASK_CONTROL_DOWN : 0)
                                | (GetKeyState((int)VIRTUAL_KEY.VK_SHIFT) < 0 ? CMIC_MASK_SHIFT_DOWN : 0)
                                ;
                            invoke.ptInvoke.X = pointScreenX;
                            invoke.ptInvoke.Y = pointScreenY;
                            invoke.nShow = (int)SHOW_WINDOW_CMD.SW_SHOWNORMAL;

                            contextMenu
                                .InvokeCommand((CMINVOKECOMMANDINFO*)&invoke)
                                .ThrowOnFailure();
                        }
                    }
                }
            }
            finally
            {
                _workingContextMenu2 = null;
                _workingContextMenu3 = null;

                if (!hMenu.IsNull)
                {
                    DestroyMenu(hMenu);
                }

                if (contextMenu is not null)
                {
                    Marshal.FinalReleaseComObject(contextMenu);
                }
            }

            return true;
        }

        static DirectoryInfo? GetParentDirectoryInfo(FileSystemInfo fileSystemInfo)
        {
            if (fileSystemInfo is FileInfo fileInfo)
            {
                return fileInfo.Directory;
            }
            else if (fileSystemInfo is DirectoryInfo directoryInfo)
            {
                return directoryInfo.Parent;
            }
            else
            {
                return null;
            }
        }

        unsafe IContextMenu CreateSingleDirectoryFileSystemItemsContextMenu(DirectoryInfo parentDir, IReadOnlyCollection<FileSystemInfo> fileSystemInfos)
        {
            IShellFolder parentFolder;
            ITEMIDLIST* parentDirItemId = null;
            object? parentFolderUnknown = null;
            object? contextMenuUnknown = null;

            ITEMIDLIST*[] items = new ITEMIDLIST*[fileSystemInfos.Count];

            try
            {
                fixed (char* displayName = parentDir.FullName)
                {
                    uint chEaten = 0;
                    uint dwAttributes = 0;
                    s_desktopFolder.ParseDisplayName(HWND.Null, null, displayName, &chEaten, &parentDirItemId, &dwAttributes);
                }

                s_desktopFolder.BindToObject(*parentDirItemId, null, typeof(IShellFolder).GUID, out parentFolderUnknown);

                parentFolder = (IShellFolder)parentFolderUnknown;

                int i = 0;
                foreach (var fileSystemInfo in fileSystemInfos)
                {
                    fixed (char* displayName = fileSystemInfo.Name)
                    fixed (ITEMIDLIST** itemPtr = &items[i++])
                    {
                        uint chEaten = 0;
                        uint dwAttributes = 0;
                        parentFolder.ParseDisplayName(HWND.Null, null, displayName, &chEaten, itemPtr, &dwAttributes);
                    }
                }

                fixed (ITEMIDLIST** itemsAryPtr = items)
                {
                    var contextMenuGuid = typeof(IContextMenu).GUID;
                    parentFolder.GetUIObjectOf(HWND.Null, (uint)items.Length, itemsAryPtr, &contextMenuGuid, null, out contextMenuUnknown);
                }

                var contextMenu = (IContextMenu)contextMenuUnknown;

                return contextMenu;
            }
            catch
            {
                if (contextMenuUnknown is not null)
                {
                    Marshal.FinalReleaseComObject(contextMenuUnknown);
                }
                throw;
            }
            finally
            {
                foreach (var itemPtr in items)
                {
                    if (itemPtr is not null)
                    {
                        Marshal.FreeCoTaskMem((IntPtr)itemPtr);
                    }
                }

                if (parentDirItemId is not null)
                {
                    Marshal.FreeCoTaskMem(new IntPtr(parentDirItemId));
                }

                if (parentFolderUnknown is not null)
                {
                    Marshal.FinalReleaseComObject(parentFolderUnknown);
                }
            }
        }

        unsafe IContextMenu CreateMultipleDirectoryFileSystemItemsContextMenu(IReadOnlyCollection<FileSystemInfo> fileSystemInfos, out bool needHandleOpenCmd, out bool ignoreEnableExtendedVerbs)
        {
            needHandleOpenCmd = true;
            ignoreEnableExtendedVerbs = true;

            object? contextMenuUnknown = null;

            ITEMIDLIST*[] items = new ITEMIDLIST*[fileSystemInfos.Count];

            try
            {
                int i = 0;
                foreach (var fileSystemInfo in fileSystemInfos)
                {
                    fixed (char* displayName = fileSystemInfo.FullName)
                    fixed (ITEMIDLIST** itemPtr = &items[i++])
                    {
                        SHParseDisplayName(displayName, null, itemPtr, 0, null);
                    }
                }

                fixed (ITEMIDLIST** itemsAryPtr = items)
                {
                    DEFCONTEXTMENU defContextMenuItem = new DEFCONTEXTMENU();
                    //defContextMenuItem.hwnd = hWnd;
                    defContextMenuItem.psf = s_desktopFolder;
                    defContextMenuItem.cidl = (uint)items.Length;
                    defContextMenuItem.apidl = itemsAryPtr;

                    SHCreateDefaultContextMenu(defContextMenuItem, typeof(IContextMenu).GUID, out contextMenuUnknown);
                }

                var contextMenu = (IContextMenu)contextMenuUnknown;

                return contextMenu;
            }
            catch
            {
                if (contextMenuUnknown is not null)
                {
                    Marshal.FinalReleaseComObject(contextMenuUnknown);
                }
                throw;
            }
            finally
            {
                foreach (var itemPtr in items)
                {
                    if (itemPtr is not null)
                    {
                        Marshal.FreeCoTaskMem((IntPtr)itemPtr);
                    }
                }
            }
        }

        unsafe IContextMenu CreateMultipleDirectoryFileSystemItemsContextMenu2(IReadOnlyCollection<FileSystemInfo> fileSystemInfos, out bool needHandleOpenCmd, out bool ignoreEnableExtendedVerbs)
        {
            needHandleOpenCmd = true;
            ignoreEnableExtendedVerbs = true;

            ITEMIDLIST*[] items = new ITEMIDLIST*[fileSystemInfos.Count];

            try
            {
                int i = 0;
                foreach (var fileSystemInfo in fileSystemInfos)
                {
                    fixed (char* displayName = fileSystemInfo.FullName)
                    fixed (ITEMIDLIST** itemPtr = &items[i++])
                    {
                        SHParseDisplayName(displayName, null, itemPtr, 0, null);
                    }
                }

                fixed (ITEMIDLIST** itemsAryPtr = items)
                {
                    
                    RegOpenKeyEx(new SafeRegistryHandle(HKEY.HKEY_CLASSES_ROOT.Value, false), "*", 0, REG_SAM_FLAGS.KEY_READ | REG_SAM_FLAGS.KEY_ENUMERATE_SUB_KEYS | REG_SAM_FLAGS.KEY_QUERY_VALUE, out var asta);
                    RegOpenKeyEx(new SafeRegistryHandle(HKEY.HKEY_CLASSES_ROOT.Value, false), ".txt", 0, REG_SAM_FLAGS.KEY_READ | REG_SAM_FLAGS.KEY_ENUMERATE_SUB_KEYS | REG_SAM_FLAGS.KEY_QUERY_VALUE, out var dotTxt);
                    RegOpenKeyEx(new SafeRegistryHandle(HKEY.HKEY_CLASSES_ROOT.Value, false), "txtfilelegacy", 0, REG_SAM_FLAGS.KEY_READ | REG_SAM_FLAGS.KEY_ENUMERATE_SUB_KEYS | REG_SAM_FLAGS.KEY_QUERY_VALUE, out var textfile);

                    var hkeys = new HKEY[]
                    {
                        new HKEY(asta.DangerousGetHandle()),
                        new HKEY(dotTxt.DangerousGetHandle()),
                        new HKEY(textfile.DangerousGetHandle()),
                    };

                    CDefFolderMenu_Create2(default, HWND.Null, (uint)items.Length, itemsAryPtr, s_desktopFolder, new LPFNDFMCALLBACK(FndFmCallback), default, out var contextMenu);

                    GC.KeepAlive(asta);
                    GC.KeepAlive(dotTxt);
                    GC.KeepAlive(textfile);

                    return contextMenu;
                }
            }
            finally
            {
                foreach (var itemPtr in items)
                {
                    if (itemPtr is not null)
                    {
                        Marshal.FreeCoTaskMem((IntPtr)itemPtr);
                    }
                }
            }
        }

        unsafe HRESULT FndFmCallback(IShellFolder psf, HWND hwnd, Windows.Win32.System.Com.IDataObject pdtobj, uint uMsg, WPARAM wParam, LPARAM lParam)
        {
            switch ((DFM_MESSAGE_ID)uMsg)
            {
                case DFM_MESSAGE_ID.DFM_MERGECONTEXTMENU: return HRESULT.S_OK;
                case DFM_MESSAGE_ID.DFM_INVOKECOMMAND: return HRESULT.S_FALSE;
                case DFM_MESSAGE_ID.DFM_GETDEFSTATICID: return HRESULT.S_FALSE;
                default: return HRESULT.E_NOTIMPL;
            }
        }

        static LRESULT InitWndProc(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
        {
            if (constructingMenuHostWindow is null)
            {
                Debug.Fail(null);
                return DefWindowProc(hWnd, msg, wParam, lParam);
            }
            else
            {
                Debug.Assert(constructingMenuHostWindow._hWnd.IsNull);
                SetWindowLongPtr(hWnd, WINDOW_LONG_PTR_INDEX.GWLP_WNDPROC, Marshal.GetFunctionPointerForDelegate(new WNDPROC(constructingMenuHostWindow.WndProc)));
                constructingMenuHostWindow._hWnd = hWnd;
                return (LRESULT)constructingMenuHostWindow.WndProc(hWnd, msg, wParam, lParam);
            }
        }

        unsafe LRESULT WndProc(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
        {
            if (_workingContextMenu3 is { } contextMenu3)
            {
                LRESULT lResult;

                if (contextMenu3.HandleMenuMsg2(msg, wParam, lParam, &lResult).Succeeded)
                {
                    return lResult;
                }
            }

            if (_workingContextMenu2 is { } contextMenu2)
            {
                switch (msg)
                {
                    case WM_INITMENUPOPUP:
                        if (contextMenu2.HandleMenuMsg(msg, wParam, lParam).Succeeded)
                        {
                            return (LRESULT)0;
                        }
                        break;
                    case WM_MEASUREITEM:
                    case WM_DRAWITEM:
                        if (contextMenu2.HandleMenuMsg(msg, wParam, lParam).Succeeded)
                        {
                            return (LRESULT)BOOL.TRUE.Value;
                        }
                        break;
                }
            }

            return DefWindowProc(hWnd, msg, wParam, lParam);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_hWnd.IsNull)
            {
                if (!DestroyWindow(_hWnd))
                {
                    Debug.Fail(null);
                    PostMessage(_hWnd, WM_CLOSE, default, default);
                }

                _hWnd = default;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}