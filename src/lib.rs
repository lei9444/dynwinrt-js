#![deny(clippy::all)]

#[cfg(target_os = "windows")]
mod winrt_service;

mod astr;

use dynwinrt;
use napi::{
  bindgen_prelude::{JsObjectValue, Promise, PromiseRaw, Result},
  Env, Error, Status,
};
use napi_derive::napi;

#[napi]
pub enum WinRTType {
  I32,
  Object,
  HString,
  HResult,
}

impl Into<dynwinrt::WinRTType> for WinRTType {
  fn into(self) -> dynwinrt::WinRTType {
    match self {
      WinRTType::I32 => dynwinrt::WinRTType::I32,
      WinRTType::Object => dynwinrt::WinRTType::Object,
      WinRTType::HString => dynwinrt::WinRTType::HString,
      WinRTType::HResult => dynwinrt::WinRTType::HResult,
    }
  }
}

#[napi]
pub async fn http_client_get_async(uri: String) -> Result<String> {
  let result = dynwinrt::http_get_string(&uri).await;
  result.map_err(|e| Error::from_reason(e.message()))
}

#[napi]
pub struct NApiVTableSignature(dynwinrt::VTableSignature);

#[napi]
impl NApiVTableSignature {
  #[napi(constructor)]
  pub fn new() -> Self {
    let iface_sig = dynwinrt::VTableSignature::new();
    NApiVTableSignature(iface_sig)
  }
}

use windows::{
  core::{AgileReference, IUnknown, Interface, HSTRING},
  Foundation::Uri,
  Web::Http::HttpProgress,
};
use windows_future::{IAsyncActionWithProgress, IAsyncOperation, IAsyncOperationWithProgress};

#[napi]
pub fn http_client_get_sync(uri: String) -> Result<ComObj> {
  use windows::Foundation::Uri;
  use windows::Web::Http::HttpClient;
  let uri = Uri::CreateUri(&HSTRING::from(uri))
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;
  let http_client =
    HttpClient::new().map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;
  let op = http_client
    .GetStringAsync(&uri)
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;
  let r = op.as_raw();
  let ukn = unsafe { IUnknown::from_raw_borrowed(&r) }.unwrap();
  Ok(ComObj(ukn.clone()))
}

#[napi]
pub async fn async_progress_hstring_to_promise_string(obj: &ComObj) -> Result<String> {
  let async_op: IAsyncOperationWithProgress<HSTRING, HttpProgress> = obj
    .0
    .cast()
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;

  let r = async_op
    .await
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;

  Ok(r.to_string())
}

#[napi]
pub fn use_dynwinrt_add(a: f64, b: f64) -> napi::Result<f64> {
  let result = dynwinrt::export_add(a, &b);
  Ok(result)
}

#[napi]
struct ComUri(Uri);

#[napi]
pub struct ComObj(IUnknown);

unsafe impl Send for ComObj {}
unsafe impl Sync for ComObj {}

impl Clone for ComObj {
  fn clone(&self) -> Self {
    ComObj(self.0.clone())
  }
}

#[napi]
pub fn from_async_string_with_http_progress(o: &ComObj) -> napi::Result<()> {
  Ok(())
}

#[napi]
struct ComObjWrapper(*mut std::ffi::c_void);

impl ComObjWrapper {
  pub fn new(ptr: *mut std::ffi::c_void) -> Self {
    ComObjWrapper(ptr)
  }
}

impl Drop for ComObjWrapper {
  fn drop(&mut self) {
    unsafe {
      if !self.0.is_null() {
        let _ = windows::core::IUnknown::from_raw(self.0);
      }
    }
  }
}

#[napi]
struct DynWinRTVTable(dynwinrt::VTableSignature);

#[napi]
impl ComUri {
  #[napi]
  pub fn createTestUri(uri_str: String) -> napi::Result<ComUri> {
    #[cfg(target_os = "windows")]
    {
      use windows::{core::HSTRING, Foundation::Uri, Win32::System::Com};
      let hstring = HSTRING::from(uri_str);

      let uri = Uri::CreateUri(&hstring)
        .map_err(|e| napi::Error::from_reason(format!("Failed to create URI: {:?}", e)))?;
      Ok(ComUri(uri))
    }

    #[cfg(not(target_os = "windows"))]
    {
      Err(napi::Error::from_reason(
        "This function is only available on Windows",
      ))
    }
  }
}

#[napi]
pub fn get_uri_vtable() -> DynWinRTVTable {
  DynWinRTVTable(dynwinrt::uri_vtable())
}

#[napi]
pub fn callMethod(vtable: &DynWinRTVTable, index: i32, obj: &ComUri) -> String {
  let m = &vtable.0.methods[index as usize];
  let result = m.call_dynamic(obj.0.as_raw(), &[]).unwrap();
  let str = result[0].as_hstring().unwrap().to_string();
  str
}

#[napi]
pub fn get_computer_name() -> napi::Result<String> {
  #[cfg(target_os = "windows")]
  {
    use windows::core::PWSTR;
    use windows::Win32::System::WindowsProgramming::GetComputerNameW;

    let mut buffer = [0u16; 256];
    let mut size = buffer.len() as u32;

    unsafe {
      if GetComputerNameW(Some(PWSTR(buffer.as_mut_ptr())), &mut size).is_ok() {
        let name = String::from_utf16_lossy(&buffer[..size as usize]);
        Ok(name)
      } else {
        Err(napi::Error::from_reason("Failed to get computer name"))
      }
    }
  }

  #[cfg(not(target_os = "windows"))]
  {
    Err(napi::Error::from_reason(
      "This function is only available on Windows",
    ))
  }
}

#[napi]
pub fn get_windows_directory() -> napi::Result<String> {
  #[cfg(target_os = "windows")]
  {
    use windows::Win32::System::SystemInformation::GetWindowsDirectoryW;

    let mut buffer = [0u16; 260]; // MAX_PATH

    unsafe {
      let len = GetWindowsDirectoryW(Some(&mut buffer));
      if len > 0 {
        let path = String::from_utf16_lossy(&buffer[..len as usize]);
        Ok(path)
      } else {
        Err(napi::Error::from_reason("Failed to get Windows directory"))
      }
    }
  }

  #[cfg(not(target_os = "windows"))]
  {
    Err(napi::Error::from_reason(
      "This function is only available on Windows",
    ))
  }
}

// Re-export WinRT service functions
#[cfg(target_os = "windows")]
pub use winrt_service::*;
