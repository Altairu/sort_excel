#[cfg(windows)]
pub mod win {
    use std::ffi::OsStr;
    use std::os::windows::ffi::OsStrExt;
    use std::ptr;
    use windows_sys::Win32::{
        Foundation::{HWND, LPARAM, LRESULT, POINT, WPARAM},
        System::LibraryLoader::GetModuleHandleW,
        UI::{
            Shell::{
                Shell_NotifyIconW, NIF_ICON, NIF_MESSAGE, NIF_TIP, NIM_ADD, NIM_DELETE,
                NOTIFYICONDATAW,
            },
            WindowsAndMessaging::{
                AppendMenuW, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu,
                DispatchMessageW, GetCursorPos, GetMessageW, LoadIconW, PostQuitMessage,
                RegisterClassW, SetForegroundWindow, TrackPopupMenu, TranslateMessage, HICON,
                HMENU, IDI_APPLICATION, MF_SEPARATOR, MF_STRING, MSG, TPM_BOTTOMALIGN,
                TPM_LEFTALIGN, TPM_RIGHTBUTTON, WM_COMMAND, WM_DESTROY, WM_LBUTTONDBLCLK,
                WM_LBUTTONUP, WM_RBUTTONUP, WM_USER, WNDCLASSW, WS_OVERLAPPED,
            },
        },
    };

    const WM_TRAY_CALLBACK: u32 = WM_USER + 101;
    const MENU_OPEN_BROWSER: usize = 1001;
    const MENU_OPEN_ADMIN: usize = 1004;
    const MENU_OPEN_FOLDER: usize = 1002;
    const MENU_QUIT: usize = 1003;

    fn to_wide(s: &str) -> Vec<u16> {
        OsStr::new(s).encode_wide().chain(Some(0)).collect()
    }

    unsafe extern "system" fn window_proc(
        hwnd: HWND,
        msg: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        match msg {
            WM_TRAY_CALLBACK => {
                let event = lparam as u32;
                if event == WM_LBUTTONUP || event == WM_LBUTTONDBLCLK {
                    let _ = open::that("http://localhost:3000");
                } else if event == WM_RBUTTONUP {
                    let mut pt = POINT { x: 0, y: 0 };
                    unsafe {
                        GetCursorPos(&mut pt);
                        SetForegroundWindow(hwnd);

                        let hmenu: HMENU = CreatePopupMenu();
                        let label = to_wide("受託確認票 サーバー (稼働中)");
                        let open_web = to_wide("使用者画面を開く (&O)");
                        let open_admin = to_wide("管理者設定画面を開く (&A)");
                        let open_dir = to_wide("Excel保存先フォルダを開く (&F)");
                        let quit_text = to_wide("終了 (&X)");

                        AppendMenuW(hmenu, MF_STRING, 0, label.as_ptr());
                        AppendMenuW(hmenu, MF_SEPARATOR, 0, ptr::null());
                        AppendMenuW(hmenu, MF_STRING, MENU_OPEN_BROWSER, open_web.as_ptr());
                        AppendMenuW(hmenu, MF_STRING, MENU_OPEN_ADMIN, open_admin.as_ptr());
                        AppendMenuW(hmenu, MF_STRING, MENU_OPEN_FOLDER, open_dir.as_ptr());
                        AppendMenuW(hmenu, MF_SEPARATOR, 0, ptr::null());
                        AppendMenuW(hmenu, MF_STRING, MENU_QUIT, quit_text.as_ptr());

                        TrackPopupMenu(
                            hmenu,
                            TPM_LEFTALIGN | TPM_BOTTOMALIGN | TPM_RIGHTBUTTON,
                            pt.x,
                            pt.y,
                            0,
                            hwnd,
                            ptr::null(),
                        );
                        DestroyMenu(hmenu);
                    }
                }
                0
            }
            WM_COMMAND => {
                match wparam {
                    MENU_OPEN_BROWSER => {
                        let _ = open::that("http://localhost:3000");
                    }
                    MENU_OPEN_ADMIN => {
                        let _ = open::that("http://localhost:3000/admin");
                    }
                    MENU_OPEN_FOLDER => {
                        let _ = open::that("exports");
                    }
                    MENU_QUIT => {
                        std::process::exit(0);
                    }
                    _ => {}
                }
                0
            }
            WM_DESTROY => {
                unsafe {
                    PostQuitMessage(0);
                }
                0
            }
            _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
        }
    }

    pub fn start_tray_loop() {
        std::thread::spawn(move || unsafe {
            let hinstance = GetModuleHandleW(ptr::null());
            let class_name = to_wide("SortExcelServerTrayClassV3");

            let mut wc: WNDCLASSW = std::mem::zeroed();
            wc.lpfnWndProc = Some(window_proc);
            wc.hInstance = hinstance;
            wc.lpszClassName = class_name.as_ptr();

            RegisterClassW(&wc);

            let hwnd = CreateWindowExW(
                0,
                class_name.as_ptr(),
                to_wide("SortExcelServerTray").as_ptr(),
                WS_OVERLAPPED,
                0,
                0,
                0,
                0,
                0,
                0,
                hinstance,
                ptr::null(),
            );

            if hwnd == 0 {
                return;
            }

            let mut hicon: HICON = LoadIconW(hinstance, 1 as _);
            if hicon == 0 {
                hicon = LoadIconW(0, IDI_APPLICATION);
            }

            let mut nid: NOTIFYICONDATAW = std::mem::zeroed();
            nid.cbSize = std::mem::size_of::<NOTIFYICONDATAW>() as u32;
            nid.hWnd = hwnd;
            nid.uID = 1;
            nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
            nid.uCallbackMessage = WM_TRAY_CALLBACK;
            nid.hIcon = hicon;

            let tip = to_wide("受託確認票 サーバー (稼働中)");
            for (i, c) in tip.iter().take(nid.szTip.len() - 1).enumerate() {
                nid.szTip[i] = *c;
            }

            let _ = Shell_NotifyIconW(NIM_ADD, &nid);

            let mut msg: MSG = std::mem::zeroed();
            while GetMessageW(&mut msg, 0, 0, 0) > 0 {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }

            Shell_NotifyIconW(NIM_DELETE, &nid);
        });
    }
}
