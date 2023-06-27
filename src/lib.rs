use std::{iter::once, os::windows::ffi::OsStrExt};
use variant_rs::dispatch::IDispatchError;
use variant_rs::dispatch::IDispatchExt;

use windows::{
    core::{ComInterface, IUnknown, PCWSTR},
    Win32::System::Com::{
        CLSIDFromProgID, CoCreateInstance, CoInitializeEx, IDispatch, CLSCTX_ALL,
        COINIT_MULTITHREADED,
    },
};
fn encode_wide(string: impl AsRef<std::ffi::OsStr>) -> Vec<u16> {
    string.as_ref().encode_wide().chain(once(0)).collect()
}
#[derive(Debug)]
pub struct DmSoft {
    pub obj: IDispatch,
}

// 大漠插件错误处理
pub enum DmError {
    IDispatchError(IDispatchError),
    DmError(i32),
    WindowsError(windows::core::Error),
}

impl std::fmt::Debug for DmError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DmError::IDispatchError(e) => write!(f, "IDispatchError: {:?}", e),
            DmError::DmError(e) => write!(f, "DmError: {:?}", e),
            DmError::WindowsError(e) => write!(f, "WindowsError: {:?}", e),
        }
    }
}

impl From<i32> for DmError {
    fn from(e: i32) -> Self {
        DmError::DmError(e)
    }
}

impl From<IDispatchError> for DmError {
    fn from(e: IDispatchError) -> Self {
        DmError::IDispatchError(e)
    }
}

impl From<windows::core::Error> for DmError {
    fn from(e: windows::core::Error) -> Self {
        DmError::WindowsError(e)
    }
}

pub type DmResult<T> = std::result::Result<T, DmError>;

impl DmSoft {
    pub fn new() -> DmResult<Self> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)?;
            let clsid = CLSIDFromProgID(PCWSTR::from_raw(encode_wide("dm.dmsoft").as_ptr()))?;
            let hr: IUnknown = CoCreateInstance(&clsid, None, CLSCTX_ALL)?;
            let obj: IDispatch = hr.cast::<IDispatch>()?;
            Ok(Self { obj })
        }
    }

    pub fn ver(&self) -> DmResult<String> {
        let res = self.obj.call("Ver", vec![])?;
        Ok(res.expect_string().to_string())
    }

    pub fn enable_pic_cache(&self) -> DmResult<i32> {
        let res = self.obj.call("EnablePicCache", vec![1.into()])?;
        Ok(res.expect_i32())
    }

    pub fn reg(&self, code: &str, ver: &str) -> DmResult<i32> {
        let res = self.obj.call("Reg", vec![code.into(), ver.into()])?;
        let res: i32 = res.expect_i32();
        Ok(res)
    }

    pub fn set_row_gap_no_dict(&self, row_gap: i64) -> DmResult<i32> {
        let res = self.obj.call("SetRowGapNoDict", vec![row_gap.into()])?;
        Ok(res.expect_i32())
    }
}
