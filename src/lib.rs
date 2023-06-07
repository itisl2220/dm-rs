use std::{
    iter::once,
    mem::ManuallyDrop,
    os::windows::{ffi::OsStrExt, prelude::OsStringExt},
    ptr,
};
use windows::{
    core::{ComInterface, IUnknown, BSTR, GUID, PCWSTR},
    Win32::System::Com::{
        CLSIDFromProgID, CoCreateInstance, CoInitializeEx, IDispatch, CLSCTX_ALL,
        COINIT_MULTITHREADED, DISPPARAMS, VARIANT,
    },
};
fn encode_wide(string: impl AsRef<std::ffi::OsStr>) -> Vec<u16> {
    string.as_ref().encode_wide().chain(once(0)).collect()
}
#[derive(Debug)]
pub struct DmSoft {
    pub obj: IDispatch,
}
struct MyDispatchDriver {
    obj: IDispatch,
}
impl MyDispatchDriver {
    pub fn new(obj: IDispatch) -> Self {
        Self { obj }
    }

    pub fn get_idof_name(&self, lpsz: PCWSTR, disp_id: &mut i32) -> () {
        return unsafe {
            let _ = self.obj.GetIDsOfNames(ptr::null(), &lpsz, 1, 0, disp_id);
        };
    }

    pub fn invoke0(&self, disp_id: i32, pvar_ret: Option<*mut VARIANT>) {
        let dis = DISPPARAMS {
            rgvarg: ptr::null_mut(),
            rgdispidNamedArgs: ptr::null_mut(),
            cArgs: 0,
            cNamedArgs: 0,
        };
        return unsafe {
            let _ = self.obj.Invoke(
                disp_id,
                &GUID::zeroed(),
                0,
                windows::Win32::System::Com::DISPATCH_METHOD,
                &dis as *const DISPPARAMS,
                pvar_ret,
                None,
                None,
            );
        };
    }

    pub fn invoke_n(
        &self,
        disp_id: i32,
        pvar_params: *mut VARIANT,
        n_params: u32,
        pvar_ret: Option<*mut VARIANT>,
    ) {
        let dis = DISPPARAMS {
            rgvarg: pvar_params,
            rgdispidNamedArgs: ptr::null_mut(),
            cArgs: n_params,
            cNamedArgs: 0,
        };
        return unsafe {
            let _ = self.obj.Invoke(
                disp_id,
                &GUID::zeroed(),
                0,
                windows::Win32::System::Com::DISPATCH_METHOD,
                &dis as *const DISPPARAMS,
                pvar_ret,
                None,
                None,
            );
        };
    }
}
impl DmSoft {
    pub fn new() -> Result<Self, anyhow::Error> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)?;
            let clsid = CLSIDFromProgID(PCWSTR::from_raw(encode_wide("dm.dmsoft").as_ptr()))?;
            let hr: IUnknown = CoCreateInstance(&clsid, None, CLSCTX_ALL)?;
            let obj: IDispatch = hr.cast::<IDispatch>()?;
            Ok(Self { obj })
        }
    }

    pub fn ver(&self) -> Result<String, anyhow::Error> {
        let mut disp_id: i32 = 0;
        let mut pvar_ret = VARIANT::default();
        let driver = MyDispatchDriver::new(self.obj.clone());
        driver.get_idof_name(PCWSTR::from_raw(encode_wide("Ver").as_ptr()), &mut disp_id);
        driver.invoke0(disp_id, Some(&mut pvar_ret));
        let pvar_ret = unsafe { &pvar_ret.Anonymous.Anonymous.Anonymous.bstrVal };
        let pvar_ret = ManuallyDrop::<BSTR>::into_inner(pvar_ret.clone());
        let pvar_ret = std::ffi::OsString::from_wide(&*pvar_ret.as_wide())
            .into_string()
            .unwrap_or(format!(""));
        Ok(pvar_ret)
    }

    pub fn reg(&self, code: String, ver: String) -> Result<i32, anyhow::Error> {
        unsafe {
            let mut pn = VARIANT::default();
            let code_v: VARIANT = VARIANT::default();
            code_v.Anonymous.Anonymous.to_owned().Anonymous.bstrVal =
                ManuallyDrop::new(BSTR::from(&code));
            let ver_v: VARIANT = VARIANT::default();
            ver_v.Anonymous.Anonymous.to_owned().Anonymous.bstrVal =
                ManuallyDrop::new(BSTR::from(&ver));
            let pn = &mut pn as *mut _ as *mut VARIANT;
            let mut disp_id: i32 = 0;
            let mut pvar_ret = VARIANT::default();
            let driver = MyDispatchDriver::new(self.obj.clone());
            driver.get_idof_name(PCWSTR::from_raw(encode_wide("Reg").as_ptr()), &mut disp_id);
            driver.invoke_n(disp_id, pn, 2, Some(&mut pvar_ret));
            let pvar_ret = pvar_ret.Anonymous.Anonymous.Anonymous.lVal;
            println!("pvar_ret: {:?}", pvar_ret);

            Ok(pvar_ret)
        }
    }

    pub fn set_row_gap_no_dict(&self, row_gap: i64) -> Result<i32, anyhow::Error> {
        unsafe {
            let v: VARIANT = VARIANT::default();
            v.Anonymous.Anonymous.to_owned().Anonymous.llVal = row_gap;
            let mut pn = [v, VARIANT::default()];
            let pn = &mut pn as *mut _ as *mut VARIANT;
            let mut disp_id: i32 = 0;
            let mut pvar_ret = VARIANT::default();
            let driver = MyDispatchDriver::new(self.obj.clone());
            driver.get_idof_name(
                PCWSTR::from_raw(encode_wide("SetRowGapNoDict").as_ptr()),
                &mut disp_id,
            );
            driver.invoke_n(disp_id, pn, 1, Some(&mut pvar_ret));
            let pvar_ret = pvar_ret.Anonymous.Anonymous.Anonymous.lVal;

            return Ok(pvar_ret);
        }
    }
}

#[cfg(test)]
mod tests {
    use windows::Win32::System::Com::{CoInitializeEx, COINIT_MULTITHREADED};

    use super::*;

    #[test]
    fn it_works() {
        unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).unwrap() };
        let obj = DmSoft::new();
        match obj {
            Ok(obj) => {
                println!("obj:{:?}", obj);
            }
            Err(e) => {
                println!("{:?}", e);
            }
        }
    }
}
