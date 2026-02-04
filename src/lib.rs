#![deny(clippy::all)]

#[cfg(target_os = "windows")]
mod winrt_service;

mod astr;

use std::rc::Rc;

use dynwinrt;
use napi::{
  bindgen_prelude::{JsObjectValue, Promise, PromiseRaw, Result},
  Env, Error, Status,
};
use napi_derive::napi;

#[napi]
pub struct DynWinRTType(dynwinrt::WinRTType);

#[napi]
impl DynWinRTType {
  #[napi]
  pub fn i32() -> Self {
    DynWinRTType(dynwinrt::WinRTType::I32)
  }

  #[napi]
  pub fn i64() -> Self {
    DynWinRTType(dynwinrt::WinRTType::I64)
  }

  #[napi]
  pub fn hstring() -> Self {
    DynWinRTType(dynwinrt::WinRTType::HString)
  }

  #[napi]
  pub fn object() -> Self {
    DynWinRTType(dynwinrt::WinRTType::Object)
  }

  #[napi]
  pub fn IAsyncOperation(iid: &WinGUID) -> Self {
    DynWinRTType(dynwinrt::WinRTType::IAsyncOperation(iid.0))
  }
}

impl Into<dynwinrt::WinRTType> for DynWinRTType {

  fn into(self) -> dynwinrt::WinRTType {
    self.0
  }
}

#[napi]
pub async fn http_client_get_async(uri: String) -> Result<String> {
  let result = dynwinrt::http_get_string(&uri).await;
  result.map_err(|e| Error::from_reason(e.message()))
}

#[napi]
pub struct NApiVTableSignature(dynwinrt::InterfaceSignature);

#[napi]
#[derive(Debug, Clone, Copy)]
pub struct WinGUID(windows::core::GUID);

#[napi]
impl WinGUID {
  #[napi]
  pub fn parse(guid_str: String) -> Self {
    let guid = windows::core::GUID::try_from(guid_str.as_str()).unwrap();
    WinGUID(guid)
  }
}

#[napi]
impl NApiVTableSignature {
  #[napi(constructor)]
  pub fn define_from_iinspectable() -> Self {
    let iface_sig = dynwinrt::InterfaceSignature::define_from_iinspectable(
      Default::default(),
      Default::default(),
    );
    NApiVTableSignature(iface_sig)
  }
}

use windows::{
  Foundation::Uri, Web::Http::{HttpClient, HttpProgress}, core::{AgileReference, HSTRING, IInspectable, IUnknown, Interface, h}
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
pub struct AsyncHSTRINGWrapper(IAsyncOperation<HSTRING>);
#[napi]
pub async fn async_hstring_to_promise_string(x: &AsyncHSTRINGWrapper) -> String {
  let r = x.0.clone().await.unwrap();
  r.to_string()
}

#[napi]
pub struct IUnknownWrapper(IUnknown);
#[napi]
pub struct AsyncUriWrapper(IAsyncOperation<HttpClient>);
// this compiles
#[napi]
pub async fn async_concrete_type_to_promise(x: &AsyncUriWrapper) -> String {
  let c = x.0.clone().await.unwrap();
  let r: IInspectable = c.cast().unwrap();
  let i = unsafe { IInspectable::from_raw(std::ptr::null_mut()) };
  // IAsyncOperation<IInspectable>::from_raw
  let i2: IAsyncOperation<IInspectable> = i.cast().unwrap();

  r.GetRuntimeClassName().unwrap().to_string()
}

#[napi]
pub struct IInspectableWrapper(IInspectable);
#[napi]
pub struct AsyncIInspectableWrapper(IAsyncOperation<IInspectable>);

// this function compiles
#[napi]
pub async fn async_inspectable_to_promise_string(x: &AsyncIInspectableWrapper) -> String {
  let r = x.0.clone().await.unwrap();
  r.GetRuntimeClassName().unwrap().to_string()
}

#[napi]
pub struct DynWinRTValue(dynwinrt::WinRTValue);

#[napi]
impl DynWinRTValue {
  #[napi]
  pub fn activation_factory(name: String) -> DynWinRTValue {
    // let f = dynwinrt::ro_get_activation_factory_2(h!("Microsoft.Windows.Storage.Pickers.FileOpenPicker")).unwrap();
    DynWinRTValue(dynwinrt::ro_get_activation_factory_2(&HSTRING::from(name)).unwrap())
  }

  #[napi]
  pub fn i64(value: i64) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::I64(value))
  }
  #[napi]
  pub fn i32(value: i32) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::I32(value))
  }
  #[napi]
  pub fn hstring(value: String) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::HString(HSTRING::from(value)))
  }
  #[napi]
  pub fn to_string(&self) -> String {
    match &self.0 {
      dynwinrt::WinRTValue::HString(s) => s.to_string(),
      dynwinrt::WinRTValue::I32(i) => i.to_string(),
      dynwinrt::WinRTValue::I64(i) => i.to_string(),
      dynwinrt::WinRTValue::Object(o) => format!("Object: {:?}", o),
      _ => "Unsupported type".to_string(),
    }
  }

  #[napi]
  pub fn call_single_out_0(
    &self,
    method_index: i32,
    return_type: &DynWinRTType
  ) -> DynWinRTValue {
    let result = self
      .0
      .call_single_out(method_index as usize, &return_type.0, &[]);
    match &result {
      Err(e) => println!("call_single_out_0 failed: {}", e.message()),
      _ => {}
    }
    DynWinRTValue(result.unwrap())
  }


  #[napi]
  pub fn call_single_out_1(
    &self,
    method_index: i32,
    return_type: &DynWinRTType,
    v1: &DynWinRTValue
  ) -> DynWinRTValue {
    let result = self
      .0
      .call_single_out(method_index as usize, &return_type.0, &[v1.0.clone()]);
    match &result {
      Err(e) => println!("call_single_out_0 failed: {}", e.message()),
      _ => {}
    }
    DynWinRTValue(result.unwrap())
  }

  #[napi]
  pub fn cast(&self, iid: &WinGUID) -> DynWinRTValue {
    let result = self.0.cast(&iid.0).unwrap();
    DynWinRTValue(result)
  }
}

#[napi]
struct WinAppSDKContext(dynwinrt::WinAppSdkContext);

#[napi]
pub fn init_winappsdk(major: u32, minor: u32) {
  WinAppSDKContext(dynwinrt::initialize_winappsdk(major, minor).unwrap());
}

#[napi]
pub async fn winrt_value_to_promise(x: &DynWinRTValue) -> DynWinRTValue {
  println!("winrt_value_to_promise called");
  let v = (&x.0).await.unwrap();
  println!("winrt_value_to_promise completed");
  DynWinRTValue(v)
}
unsafe impl Send for DynWinRTValue {}
unsafe impl Sync for DynWinRTValue {}

#[napi]
pub struct WinIIds(dynwinrt::IIds);

#[napi]
pub async fn run_test_picker() {
  dynwinrt::test_pick_open_picker_full_dynamic().await.unwrap();
}

#[napi]
impl WinIIds {
  #[napi]
  pub fn IFileOpenPickerFactoryIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::IFileOpenPickerFactory)
  }

  #[napi]
  pub fn IAsyncOperationPickFileResultIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::IAsyncOperationPickFileResult)
  }
}

// this gives a compile error of Send/Sync is required,
// which indicate that both argument and return type must be Send + Sync for napi
#[napi]
pub async fn async_inspectable_to_promise_inspectable(
  x: &AsyncIInspectableWrapper,
) -> IInspectableWrapper {
  let r = x.0.clone().await.unwrap();
  IInspectableWrapper(r)
}
// add following by pass compile error
unsafe impl Send for IInspectableWrapper {}
unsafe impl Sync for IInspectableWrapper {}

// this gives multiple compiles error:
// 1 IUnknown is not RuntimeType so IAsyncOperation<IUnknown> is not valid
// 2 even wrapped in IAsyncOperation, IAsyncOperation<IUnknown> is not Send + Sync from napi's perspective
#[napi]
pub struct IAsyncIUnknownWrapper(IAsyncOperation<IInspectable>);
// #[napi]
// pub async fn unwrap_async_iunknown2(x: &IAsyncIUnknownWrapper) -> String {
//   let async_op: IAsyncOperation<IUnknown> = x.0.cast().unwrap();
//   let r: IInspectable = async_op.await.cast().unwrap();
//   r.GetRuntimeClassName().unwrap().to_string()
// }
// fn foo(x: IAsyncOperation<IInspectable>) {}

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
struct DynWinRTVTable(dynwinrt::InterfaceSignature);

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
