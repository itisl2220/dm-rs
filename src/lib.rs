use std::{iter::once, os::windows::ffi::OsStrExt};
use variant_rs::dispatch::IDispatchError;
use variant_rs::dispatch::IDispatchExt;

use windows::core::PCWSTR;
use windows::Win32::Foundation::{HWND, RECT};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::UI::WindowsAndMessaging::GetClientRect;
use windows::{
    core::{ComInterface, IUnknown},
    Win32::System::Com::{
        CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_MULTITHREADED,
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

    pub fn set_path(&self, path: &str) -> DmResult<i32> {
        let res = self.obj.call("SetPath", vec![path.into()])?;
        Ok(res.expect_i32())
    }

    // 设置字库
    pub fn set_dict(&self, index: i64, file: &str) -> DmResult<i32> {
        let res = self.obj.call("SetDict", vec![index.into(), file.into()])?;
        Ok(res.expect_i32())
    }

    pub fn set_row_gap_no_dict(&self, row_gap: i64) -> DmResult<i32> {
        let res = self.obj.call("SetRowGapNoDict", vec![row_gap.into()])?;
        Ok(res.expect_i32())
    }

    // 设置窗口大小
    pub fn set_window_size(&self, hwnd: i64, width: i64, height: i64) -> DmResult<i32> {
        let res = self.obj.call(
            "SetWindowSize",
            vec![hwnd.into(), width.into(), height.into()],
        )?;
        Ok(res.expect_i32())
    }

    // 获取窗口客户区域大小
    pub fn get_client_size(&self, hwnd: i64) -> DmResult<(i32, i32)> {
        let mut rect = RECT::default();
        let res = unsafe { GetClientRect(HWND(hwnd as isize), &mut rect) };
        if res.0 == 0 {
            return Err(DmError::WindowsError(windows::core::Error::from_win32()));
        }
        Ok((rect.right, rect.bottom))
    }

    // 设置窗口客户区域大小
    pub fn set_client_size(&self, hwnd: i64, width: i64, height: i64) -> DmResult<i32> {
        let res = self.obj.call(
            "SetClientSize",
            vec![hwnd.into(), width.into(), height.into()],
        )?;
        Ok(res.expect_i32())
    }

    // 高级窗口绑定
    pub fn bind_window_ex(
        &self,
        hwnd: i64,
        display: &str,
        mouse: &str,
        keypad: &str,
        public: &str,
        mode: i64,
    ) -> DmResult<i32> {
        let res = self.obj.call(
            "BindWindowEx",
            vec![
                hwnd.into(),
                display.into(),
                mouse.into(),
                keypad.into(),
                public.into(),
                mode.into(),
            ],
        )?;
        Ok(res.expect_i32())
    }

    // 文字识别
    pub fn ocr(
        &self,
        x1: i64,
        y1: i64,
        x2: i64,
        y2: i64,
        color: &str,
        sim: f64,
    ) -> DmResult<String> {
        let res = self.obj.call(
            "Ocr",
            vec![
                x1.into(),
                y1.into(),
                x2.into(),
                y2.into(),
                color.into(),
                sim.into(),
            ],
        )?;
        Ok(res.expect_string().to_string())
    }

    // 高级找图
    pub fn find_pic_ex(
        &self,
        x1: i64,
        y1: i64,
        x2: i64,
        y2: i64,
        pic_name: &str,
        delta_color: &str,
        sim: f64,
        dir: i64,
    ) -> DmResult<String> {
        let res = self.obj.call(
            "FindPicEx",
            vec![
                x1.into(),
                y1.into(),
                x2.into(),
                y2.into(),
                pic_name.into(),
                delta_color.into(),
                sim.into(),
                dir.into(),
            ],
        )?;
        Ok(res.expect_string().to_string())
    }

    // 移动鼠标
    pub fn move_to(&self, x: i64, y: i64) -> DmResult<i32> {
        let res = self.obj.call("MoveTo", vec![x.into(), y.into()])?;
        Ok(res.expect_i32())
    }

    // 鼠标左键单击
    pub fn left_click(&self) -> DmResult<i32> {
        let res = self.obj.call("LeftClick", vec![])?;
        Ok(res.expect_i32())
    }

    // 查找指定区域内的颜色,颜色格式"RRGGBB-DRDGDB",注意,和按键的颜色格式相反
    pub fn find_color_e(
        &self,
        x1: i64,
        y1: i64,
        x2: i64,
        y2: i64,
        color: &str,
        sim: f64,
        dir: i64,
    ) -> DmResult<String> {
        let res = self.obj.call(
            "FindColorE",
            vec![
                x1.into(),
                y1.into(),
                x2.into(),
                y2.into(),
                color.into(),
                sim.into(),
                dir.into(),
            ],
        )?;
        Ok(res.expect_string().to_string())
    }
}
