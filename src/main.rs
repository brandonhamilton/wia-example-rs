extern crate winapi;

use std::ptr::null_mut;
use winapi::ctypes::c_void;
use winapi::shared::basetsd::UINT32;
use winapi::shared::wtypesbase::{CLSCTX_LOCAL_SERVER, OLECHAR};
use winapi::shared::wtypes::BSTR;
use winapi::um::combaseapi::{CoCreateInstance, CoUninitialize};
use winapi::um::wia::*;
use winapi::um::objbase::CoInitialize;
use winapi::um::oleauto::SysAllocStringLen;
use winapi::um::winnt::{LONG};
use winapi::shared::winerror;

fn create_bstr(s: &str) -> BSTR {
    let mut s16: Vec<u16> = Vec::with_capacity(s.len());
    for c in s.encode_utf16() {
        s16.push(c);
    }
    let len = s16.len();
    let slice: &[u16] = &s16;
    unsafe { SysAllocStringLen(slice as *const _ as *const OLECHAR, len as UINT32) }
}

fn main() {
    let mut hr = unsafe { CoInitialize(null_mut()) };
    if winerror::SUCCEEDED(hr) {

        let mut pWiaDevMgr2: *mut c_void = null_mut();
        hr = unsafe { CoCreateInstance(
            &CLSID_WiaDevMgr2,
            null_mut(),
            CLSCTX_LOCAL_SERVER,
            &IID_IWiaDevMgr2,
            &mut pWiaDevMgr2,
        ) };

        if winerror::SUCCEEDED(hr) {

            let pp_wia_dev_mgr: *mut IWiaDevMgr2 = pWiaDevMgr2 as *mut IWiaDevMgr2;

            let folder:  BSTR = create_bstr("c:\\");
            let filename: BSTR = create_bstr("Scan");

            let mut num_files: LONG = 0;
            let mut file_paths: *mut BSTR = null_mut();
            let mut p_item: *mut IWiaItem2 = null_mut();
            hr = unsafe { (*pp_wia_dev_mgr).GetImageDlg(
                0,
                null_mut(),
                null_mut(),
                folder,
                filename,
                &mut num_files,
                &mut file_paths,
                &mut p_item,
            ) };

            if winerror::SUCCEEDED(hr) {
                println!("Image successfully scanned");
            } else {    
                println!("Unable to call GetImageDlg on  IWiaDevMgr2: {:x}", hr);
            }
            unsafe { (*pp_wia_dev_mgr).Release(); }
        } else {
            println!("Unable to create IWiaDevMgr2 interface: {:x}", hr);
        }
        unsafe { CoUninitialize(); }
    } else {
        println!("Unable to initialize COM: {:x}", hr);
    }
}
