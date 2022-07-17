extern crate lazy_static;

use std::borrow::BorrowMut;
use std::env;
use std::ffi::{c_void, CString};
use std::path::Path;
use std::ptr::null_mut;
use std::sync::mpsc;
use std::time::Duration;

use clokwerk::{Scheduler, TimeUnits};
use lazy_static::lazy_static;
use tray_item::TrayItem;
use windows::core::{GUID, HSTRING, IInspectable, Interface, IUnknown, PCSTR, PCWSTR};
use windows::Data::Xml::Dom::{XmlDocument, XmlElement, XmlNodeList};
use windows::Foundation::TypedEventHandler;
use windows::UI::Notifications::{
    ToastActivatedEventArgs, ToastNotification, ToastNotificationManager,
    ToastNotificationManagerForUser, ToastTemplateType,
};
use windows::Win32::Storage::EnhancedStorage::PKEY_AppUserModel_ID;
use windows::Win32::System::Com::{
    CLSCTX_INPROC_SERVER, CoCreateInstance, CoInitialize, IPersistFile,
};
use windows::Win32::System::Registry::{
    HKEY, HKEY_LOCAL_MACHINE, KEY_SET_VALUE, REG_EXPAND_SZ, RegOpenKeyExA, RegSetKeyValueA,
};
use windows::Win32::UI::Shell::{IShellLinkA, ShellLink};
use windows::Win32::UI::Shell::PropertiesSystem::{
    InitPropVariantFromGUIDAsString, IPropertyStore,
};


enum TrayMessage {
    Quit,
}

lazy_static! {
    static ref APP_GUID: GUID = GUID::from("6825c41d-98c8-4ff6-90bc-00acf336d8f3");
}
const APP_ID_STRING: &str = "LazyEye";

unsafe fn set_registry_key(h_key: HKEY, name: &str, data: &str) {
    let sub_key = CString::new("SOFTWARE\\Classes\\AppUserModelId\\LazyEye").unwrap();
    let sub_key_ptr = sub_key.as_bytes_with_nul().as_ptr();
    let sub_key_pcstr = PCSTR(sub_key_ptr);

    let value_name = CString::new(name).unwrap();
    let value_name_ptr = value_name.as_bytes_with_nul().as_ptr();
    let value_name_pcstr = PCSTR(value_name_ptr);

    let value_data = String::from(data);
    let value_data_c_string = CString::new(value_data).unwrap();
    let value_data_c_void = value_data_c_string.as_bytes_with_nul().as_ptr() as *const c_void;

    RegSetKeyValueA(
        h_key,
        sub_key_pcstr,
        value_name_pcstr,
        REG_EXPAND_SZ.0,
        value_data_c_void,
        value_data_c_string.as_bytes_with_nul().len() as u32,
    );
}

unsafe fn create_shortcut(path: String) {
    let exe_path = CString::new(env::current_exe().unwrap().to_str().unwrap()).unwrap();
    let exe_path_ptr = exe_path.as_bytes_with_nul().as_ptr();
    let exe_path_pcstr = PCSTR(exe_path_ptr);

    let shell_link: IShellLinkA =
        CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER).unwrap();

    shell_link.SetPath(exe_path_pcstr).unwrap();

    let property_store: IPropertyStore = shell_link.cast().unwrap();

    property_store
        .SetValue(
            &PKEY_AppUserModel_ID,
            &InitPropVariantFromGUIDAsString(&*APP_GUID).unwrap(),
        )
        .unwrap();
    property_store.Commit().unwrap();

    let persist_file: IPersistFile = property_store.cast().unwrap();
    let mut shortcut_path_u16: Vec<u16> = path.encode_utf16().collect();
    shortcut_path_u16.push(0);
    let shortcut_path_pcwstr = PCWSTR(shortcut_path_u16.as_ptr());
    persist_file.Save(shortcut_path_pcwstr, true).unwrap();
}

fn create_shortcut_and_registry_data_if_not_exists() {
    let mut shortcut_path = env::var("APPDATA").unwrap();
    shortcut_path.push_str("\\Microsoft\\Windows\\Start Menu\\Programs\\Lazy Eye.lnk");
    let shortcut_exists = Path::new(&shortcut_path).exists();
    if shortcut_exists {
        return;
    } else {
        unsafe {
            CoInitialize(null_mut()).unwrap();

            let mut h_key: HKEY = HKEY::default();
            RegOpenKeyExA(
                HKEY_LOCAL_MACHINE,
                None,
                0,
                KEY_SET_VALUE,
                h_key.borrow_mut(),
            );
            set_registry_key(h_key, "DisplayName", "Lazy Eye");
            create_shortcut(shortcut_path);
        }
    }
}

fn send_notification(manager: &ToastNotificationManagerForUser) {
    let content: XmlDocument =
        ToastNotificationManager::GetTemplateContent(ToastTemplateType::ToastText01).unwrap();

    let toast_element: XmlElement = content.DocumentElement().unwrap();

    toast_element.SetAttribute(HSTRING::from("scenario"), HSTRING::from("reminder")).unwrap();

    let start_timer_action_element: XmlElement =
        content.CreateElement(HSTRING::from("action")).unwrap();
    start_timer_action_element.SetAttribute(HSTRING::from("content"), HSTRING::from("Start Timer")).unwrap();
    start_timer_action_element
        .SetAttribute(HSTRING::from("arguments"), HSTRING::from("start_timer")).unwrap();
    start_timer_action_element.SetAttribute(HSTRING::from("type"), HSTRING::from("start_timer")).unwrap();

    let dismiss_action_element: XmlElement =
        content.CreateElement(HSTRING::from("action")).unwrap();
    dismiss_action_element.SetAttribute(HSTRING::from("content"), HSTRING::from("Dismiss")).unwrap();
    dismiss_action_element.SetAttribute(HSTRING::from("arguments"), HSTRING::from("dismiss")).unwrap();
    dismiss_action_element.SetAttribute(HSTRING::from("type"), HSTRING::from("dismiss")).unwrap();

    let actions_element: XmlElement = content.CreateElement(HSTRING::from("actions")).unwrap();
    actions_element.AppendChild(start_timer_action_element).unwrap();
    actions_element.AppendChild(dismiss_action_element).unwrap();

    toast_element.AppendChild(actions_element).unwrap();

    let node_list: XmlNodeList = content.GetElementsByTagName(HSTRING::from("text")).unwrap();
    for i in 0..node_list.Length().unwrap() {
        node_list.GetAt(i).unwrap().AppendChild(
            content
                .CreateTextNode(HSTRING::from("Time to let your eyes rest."))
                .unwrap(),
        ).unwrap();
    }

    let notification: ToastNotification =
        ToastNotification::CreateToastNotification(content).unwrap();
    notification.Activated(TypedEventHandler::new(
        |_sender, result: &Option<IInspectable>| {
            let t: &IUnknown = &result.as_ref().unwrap().0;
            let t: ToastActivatedEventArgs = t.cast().unwrap();
            let argument = t.Arguments().unwrap().to_string_lossy();
            let argument_str: &str = argument.as_str();

            match argument_str {
                "start_timer" => {
                    // TODO: Implement Timer Functionality.
                }
                _ => {}
            }

            Ok(())
        },
    )).unwrap();
    let notifier = manager
        .CreateToastNotifierWithId(HSTRING::from(APP_ID_STRING))
        .unwrap();
    notifier.Show(notification).unwrap();
}

fn main() {
    create_shortcut_and_registry_data_if_not_exists();

    let mut tray = TrayItem::new("Lazy Eye", "lazy-eye-icon").unwrap();
    let (tx, rx) = mpsc::channel();

    tray.add_label("Lazy Eye").unwrap();
    tray.add_menu_item("Quit", move || {
        tx.send(TrayMessage::Quit).unwrap();
    })
        .unwrap();

    let manager: ToastNotificationManagerForUser = ToastNotificationManager::GetDefault().unwrap();

    let mut scheduler = Scheduler::new();
    scheduler
        .every(1.minutes())
        .run(move || send_notification(&manager));
    let thread_handle = scheduler.watch_thread(Duration::from_millis(100));
    loop {
        let received = rx.recv().unwrap();
        match received {
            TrayMessage::Quit => {
                thread_handle.stop();
                break;
            }
        }
    }
}
