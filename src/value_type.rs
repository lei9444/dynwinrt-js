use std::sync::Arc;

use dynwinrt::registry::{PrimitiveType, TypeHandle, TypeRegistry, ValueTypeData};
use napi::Result;
use napi_derive::napi;

#[napi]
pub fn ro_initialize_mta() -> Result<()> {
  unsafe {
    windows::Win32::System::WinRT::RoInitialize(
      windows::Win32::System::WinRT::RO_INIT_MULTITHREADED,
    )
    .map_err(|e| napi::Error::from_reason(format!("{}", e)))?;
  }
  Ok(())
}

#[napi]
pub struct JsTypeRegistry {
  inner: Arc<TypeRegistry>,
}

#[napi]
impl JsTypeRegistry {
  #[napi(constructor)]
  pub fn new() -> Self {
    Self {
      inner: TypeRegistry::new(),
    }
  }

  #[napi]
  pub fn bool_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::Bool))
  }

  #[napi]
  pub fn u8_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::U8))
  }

  #[napi]
  pub fn i16_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::I16))
  }

  #[napi]
  pub fn u16_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::U16))
  }

  #[napi]
  pub fn i32_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::I32))
  }

  #[napi]
  pub fn u32_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::U32))
  }

  #[napi]
  pub fn i64_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::I64))
  }

  #[napi]
  pub fn u64_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::U64))
  }

  #[napi]
  pub fn f32_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::F32))
  }

  #[napi]
  pub fn f64_type(&self) -> JsTypeHandle {
    JsTypeHandle(self.inner.primitive(PrimitiveType::F64))
  }

  #[napi]
  pub fn define_struct(&self, fields: Vec<&JsTypeHandle>) -> JsTypeHandle {
    let handles: Vec<TypeHandle> = fields.iter().map(|h| h.0.clone()).collect();
    JsTypeHandle(self.inner.define_struct(&handles))
  }
}

#[napi]
pub struct JsTypeHandle(TypeHandle);

#[napi]
impl JsTypeHandle {
  #[napi]
  pub fn size_of(&self) -> u32 {
    self.0.size_of() as u32
  }

  #[napi]
  pub fn align_of(&self) -> u32 {
    self.0.align_of() as u32
  }

  #[napi]
  pub fn field_count(&self) -> u32 {
    self.0.field_count() as u32
  }

  #[napi]
  pub fn field_offset(&self, index: u32) -> u32 {
    self.0.field_offset(index as usize) as u32
  }

  #[napi]
  pub fn field_type(&self, index: u32) -> JsTypeHandle {
    JsTypeHandle(self.0.field_type(index as usize))
  }

  #[napi]
  pub fn default_value(&self) -> JsValueTypeData {
    JsValueTypeData(self.0.default_value())
  }
}

#[napi]
pub struct JsValueTypeData(ValueTypeData);

unsafe impl Send for JsValueTypeData {}
unsafe impl Sync for JsValueTypeData {}

#[napi]
impl JsValueTypeData {
  // Typed getters for JS (no generics in napi)

  #[napi]
  pub fn get_f32(&self, index: u32) -> f64 {
    self.0.get_field::<f32>(index as usize) as f64
  }

  #[napi]
  pub fn get_f64(&self, index: u32) -> f64 {
    self.0.get_field::<f64>(index as usize)
  }

  #[napi]
  pub fn get_i32(&self, index: u32) -> i32 {
    self.0.get_field::<i32>(index as usize)
  }

  #[napi]
  pub fn get_i64(&self, index: u32) -> i64 {
    self.0.get_field::<i64>(index as usize)
  }

  #[napi]
  pub fn get_u8(&self, index: u32) -> u32 {
    self.0.get_field::<u8>(index as usize) as u32
  }

  #[napi]
  pub fn get_u16(&self, index: u32) -> u32 {
    self.0.get_field::<u16>(index as usize) as u32
  }

  #[napi]
  pub fn get_u32(&self, index: u32) -> u32 {
    self.0.get_field::<u32>(index as usize)
  }

  #[napi]
  pub fn get_u64(&self, index: u32) -> i64 {
    self.0.get_field::<u64>(index as usize) as i64
  }

  // Typed setters

  #[napi]
  pub fn set_f32(&mut self, index: u32, value: f64) {
    self.0.set_field(index as usize, value as f32);
  }

  #[napi]
  pub fn set_f64(&mut self, index: u32, value: f64) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn set_i32(&mut self, index: u32, value: i32) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn set_i64(&mut self, index: u32, value: i64) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn set_u8(&mut self, index: u32, value: u32) {
    self.0.set_field(index as usize, value as u8);
  }

  #[napi]
  pub fn set_u16(&mut self, index: u32, value: u32) {
    self.0.set_field(index as usize, value as u16);
  }

  #[napi]
  pub fn set_u32(&mut self, index: u32, value: u32) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn set_u64(&mut self, index: u32, value: i64) {
    self.0.set_field(index as usize, value as u64);
  }

  /// Call a COM method that takes this struct by value and returns an Object.
  /// ABI pattern: HRESULT Method(this_ptr, struct_by_value, *out_ptr)
  #[napi]
  pub fn call_method(
    &self,
    obj: &crate::DynWinRTValue,
    method_index: i32,
  ) -> Result<crate::DynWinRTValue> {
    let obj_raw = obj.as_raw_ptr();
    let result = self
      .0
      .call_method_struct_to_object(obj_raw, method_index as usize)
      .map_err(|e| napi::Error::from_reason(format!("{}", e)))?;
    Ok(crate::DynWinRTValue::from_iunknown(result))
  }
}

// Ad hoc helpers: extract Geopoint.Position fields via static projection.
// These bypass the need for struct-return support, used only for testing.

#[napi]
pub fn geopoint_get_latitude(obj: &crate::DynWinRTValue) -> Result<f64> {
  use windows::core::Interface;
  let raw = obj.as_raw_ptr();
  let geopoint =
    unsafe { windows::Devices::Geolocation::Geopoint::from_raw_borrowed(&raw) }
      .ok_or_else(|| napi::Error::from_reason("null pointer"))?
      .clone();
  let pos = geopoint
    .Position()
    .map_err(|e| napi::Error::from_reason(format!("{}", e)))?;
  Ok(pos.Latitude)
}

#[napi]
pub fn geopoint_get_longitude(obj: &crate::DynWinRTValue) -> Result<f64> {
  use windows::core::Interface;
  let raw = obj.as_raw_ptr();
  let geopoint =
    unsafe { windows::Devices::Geolocation::Geopoint::from_raw_borrowed(&raw) }
      .ok_or_else(|| napi::Error::from_reason("null pointer"))?
      .clone();
  let pos = geopoint
    .Position()
    .map_err(|e| napi::Error::from_reason(format!("{}", e)))?;
  Ok(pos.Longitude)
}
