use std::{iter::once, ops::DerefMut, os::windows::ffi::OsStrExt, ptr};
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
    pub fn new(obj: &IDispatch) -> Self {
        Self { obj: obj.clone() }
    }

    pub fn get_idof_name(&self, lpsz: &str, disp_id: &mut i32) -> () {
        let lpsz = PCWSTR::from_raw(encode_wide(lpsz).as_ptr());
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
                &dis,
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
        unsafe {
            let mut disp_id: i32 = 0;
            let mut pvar_ret = VARIANT::default();
            let driver = MyDispatchDriver::new(&self.obj);
            driver.get_idof_name("Ver", &mut disp_id);
            driver.invoke0(disp_id, Some(&mut pvar_ret));
            let result = &pvar_ret.Anonymous.Anonymous.Anonymous.bstrVal;
            Ok(result.to_string())
        }
    }

    pub fn enable_pic_cache(&self) -> Result<String, anyhow::Error> {
        unsafe {
            let mut disp_id: i32 = 0;
            let mut pvar_ret = VARIANT::default();

            let mut pn: [VARIANT; 1] = [VARIANT::default()];
            pn[0].Anonymous.Anonymous.deref_mut().Anonymous.llVal = 1;

            let driver = MyDispatchDriver::new(&self.obj);
            driver.get_idof_name("EnablePicCache", &mut disp_id);
            driver.invoke_n(disp_id, pn.as_mut_ptr(), 1, Some(&mut pvar_ret));
            let result = &pvar_ret.Anonymous.Anonymous.Anonymous.bstrVal;
            Ok(result.to_string())
        }
    }

    pub fn reg(&self, code: &str, ver: &str) -> Result<i32, anyhow::Error> {
        unsafe {
            let mut pn: [VARIANT; 2] = [VARIANT::default(), VARIANT::default()];
            *pn[0].Anonymous.Anonymous.deref_mut().Anonymous.bstrVal = BSTR::from(code);
            *pn[0].Anonymous.Anonymous.deref_mut().Anonymous.bstrVal = BSTR::from(ver);
            let mut disp_id: i32 = -1;
            let mut pvar_ret = VARIANT::default();
            let driver = MyDispatchDriver::new(&self.obj);
            driver.get_idof_name("Reg", &mut disp_id);
            driver.invoke_n(disp_id, pn.as_mut_ptr(), 2, Some(&mut pvar_ret));
            Ok(pvar_ret.Anonymous.Anonymous.Anonymous.lVal)
        }
    }

    pub fn set_row_gap_no_dict(&self, row_gap: i64) -> Result<i32, anyhow::Error> {
        unsafe {
            let mut pn = [VARIANT::default()];
            (*pn[0].Anonymous.Anonymous).Anonymous.llVal = row_gap;
            let mut disp_id: i32 = 0;
            let mut pvar_ret = VARIANT::default();
            let driver = MyDispatchDriver::new(&self.obj);
            driver.get_idof_name("SetRowGapNoDict", &mut disp_id);
            driver.invoke_n(disp_id, pn.as_mut_ptr(), 1, Some(&mut pvar_ret));
            let pvar_ret = pvar_ret.Anonymous.Anonymous.Anonymous.lVal;

            return Ok(pvar_ret);
        }
    }
}
