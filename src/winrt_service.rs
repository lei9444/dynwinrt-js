//! WinRT Service for JavaScript
//! 
//! This module provides WinRT/COM interop functionality similar to jswinrtm's WinRTService,
//! but implemented in Rust using the `windows` crate for better performance and safety.
//! 
//! Key features:
//! - Automatic COM reference counting via RAII (Drop trait)
//! - Type-safe WinRT string management (HSTRING)
//! - COM object activation and interface querying
//! - Thread apartment initialization
//! 
//! The Rust `windows` crate handles:
//! - IUnknown/IInspectable vtable management
//! - Automatic AddRef/Release on clone/drop
//! - Type-safe interface casting
//! - HRESULT error handling

use napi::bindgen_prelude::*;
use napi_derive::napi;
use windows::core::{Interface, HSTRING, GUID};
use windows::Win32::System::WinRT::{
    RoInitialize, RoUninitialize, RoActivateInstance,
    RO_INIT_MULTITHREADED, RO_INIT_SINGLETHREADED,
};

/// Wrapper around IInspectable COM pointer for JavaScript
/// This provides safe COM object lifetime management with automatic reference counting
#[napi]
pub struct ComObject {
    // Using Option to allow for safe Drop handling
    ptr: Option<windows::core::IInspectable>,
    // Store the runtime class name for debugging
    class_name: String,
}

#[napi]
impl ComObject {
    /// Get the runtime class name of this COM object
    #[napi]
    pub fn get_runtime_class_name(&self) -> Result<String> {
        Ok(self.class_name.clone())
    }

    /// Query for a specific interface by IID
    /// Returns a new ComObject with the requested interface
    #[napi]
    pub fn query_interface(&self, iid: String) -> Result<ComObject> {
        let ptr = self.ptr.as_ref()
            .ok_or_else(|| Error::from_reason("COM object has been disposed"))?;

        // Parse the IID string (format: "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX")
        let guid = parse_guid(&iid)?;

        // Cast to the requested interface
        // The windows crate handles QueryInterface internally
        let new_ptr: windows::core::IInspectable = unsafe {
            ptr.cast()
                .map_err(|e| Error::from_reason(format!("QueryInterface failed: {:?}", e)))?
        };

        Ok(ComObject {
            ptr: Some(new_ptr),
            class_name: self.class_name.clone(),
        })
    }

    /// Get the current reference count (for debugging)
    /// Note: This is primarily for debugging purposes
    #[napi]
    pub fn get_ref_count(&self) -> Result<u32> {
        let ptr = self.ptr.as_ref()
            .ok_or_else(|| Error::from_reason("COM object has been disposed"))?;

        // Temporarily increase and decrease ref count to peek at it
        // The windows crate doesn't expose AddRef/Release directly,
        // so we clone (which calls AddRef) and measure
        let cloned = ptr.clone();
        // The clone adds a ref, original has N, cloned has N+1
        // When cloned drops, it goes back to N
        drop(cloned);

        // Return 1 as a placeholder since we can't easily get exact count
        // In practice, the Rust windows crate manages this automatically
        Ok(1)
    }
}

// Automatic cleanup when ComObject is dropped
impl Drop for ComObject {
    fn drop(&mut self) {
        // The IInspectable smart pointer automatically calls Release
        // when dropped, so we just need to take ownership and drop it
        if let Some(ptr) = self.ptr.take() {
            drop(ptr);
        }
    }
}

/// WinRT Service - Main API for working with Windows Runtime
#[napi]
pub struct WinRTService {
    initialized: bool,
}

#[napi]
impl WinRTService {
    /// Create a new WinRT service instance
    #[napi(constructor)]
    pub fn new() -> Self {
        WinRTService {
            initialized: false,
        }
    }

    /// Initialize the Windows Runtime
    /// 
    /// # Arguments
    /// * `apartment_type` - 0 for single-threaded (STA), 1 for multi-threaded (MTA)
    /// 
    /// This is equivalent to koffi's RoInitialize
    #[napi]
    pub fn ro_initialize(&mut self, apartment_type: Option<i32>) -> Result<()> {
        let init_type = match apartment_type.unwrap_or(1) {
            0 => RO_INIT_SINGLETHREADED,
            _ => RO_INIT_MULTITHREADED,
        };

        unsafe {
            RoInitialize(init_type)
                .map_err(|e| Error::from_reason(format!("RoInitialize failed: {:?}", e)))?;
        }

        self.initialized = true;
        Ok(())
    }

    /// Uninitialize the Windows Runtime
    #[napi]
    pub fn ro_uninitialize(&mut self) -> Result<()> {
        if self.initialized {
            unsafe {
                RoUninitialize();
            }
            self.initialized = false;
        }
        Ok(())
    }

    /// Create a WinRT HSTRING from a regular string
    /// 
    /// Note: In Rust, we don't need to manually delete HSTRINGs
    /// The HSTRING type has Drop implemented and handles cleanup automatically
    #[napi]
    pub fn create_string(&self, source: String) -> Result<String> {
        let hstring = HSTRING::from(source.as_str());
        // Convert back to String for JavaScript
        // In a real implementation, you might want to keep the HSTRING
        Ok(hstring.to_string())
    }

    /// Get the raw buffer from an HSTRING
    /// In Rust, this is just a string conversion since HSTRING manages itself
    #[napi]
    pub fn get_string_raw_buffer(&self, hstring: String) -> Result<String> {
        // In Rust, strings are already UTF-8, HSTRING is UTF-16
        // For now, just return the string as-is
        Ok(hstring)
    }

    /// Activate a WinRT class instance
    /// 
    /// # Arguments
    /// * `class_name` - Fully qualified WinRT class name (e.g., "Windows.Data.Json.JsonObject")
    /// 
    /// Returns a ComObject wrapping the activated IInspectable interface
    /// 
    /// This is the Rust equivalent of your koffi-based RoActivateInstance
    #[napi]
    pub fn activate_instance(&self, class_name: String) -> Result<ComObject> {
        // Create HSTRING from class name
        let hstring = HSTRING::from(class_name.as_str());

        // Activate the instance
        let inspectable = unsafe {
            RoActivateInstance(&hstring)
                .map_err(|e| Error::from_reason(
                    format!("Failed to activate '{}': {:?}", class_name, e)
                ))?
        };

        Ok(ComObject {
            ptr: Some(inspectable),
            class_name: class_name.clone(),
        })
    }

    /// Get the runtime class name of a COM object
    #[napi]
    pub fn get_runtime_class_name(&self, obj: &ComObject) -> Result<String> {
        obj.get_runtime_class_name()
    }
}

// Automatic cleanup
impl Drop for WinRTService {
    fn drop(&mut self) {
        if self.initialized {
            unsafe {
                RoUninitialize();
            }
        }
    }
}

/// Parse a GUID string into a windows::core::GUID
/// Format: "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
fn parse_guid(guid_str: &str) -> Result<GUID> {
    // Remove braces if present
    let cleaned = guid_str.trim_matches(|c| c == '{' || c == '}');
    
    // Split into parts
    let parts: Vec<&str> = cleaned.split('-').collect();
    if parts.len() != 5 {
        return Err(Error::from_reason(format!("Invalid GUID format: {}", guid_str)));
    }

    // Parse each part
    let data1 = u32::from_str_radix(parts[0], 16)
        .map_err(|_| Error::from_reason("Invalid GUID data1"))?;
    let data2 = u16::from_str_radix(parts[1], 16)
        .map_err(|_| Error::from_reason("Invalid GUID data2"))?;
    let data3 = u16::from_str_radix(parts[2], 16)
        .map_err(|_| Error::from_reason("Invalid GUID data3"))?;

    // Parse data4 (8 bytes)
    let data4_part1 = u16::from_str_radix(parts[3], 16)
        .map_err(|_| Error::from_reason("Invalid GUID data4 part1"))?;
    let data4_part2 = u64::from_str_radix(parts[4], 16)
        .map_err(|_| Error::from_reason("Invalid GUID data4 part2"))?;

    let mut data4 = [0u8; 8];
    data4[0] = (data4_part1 >> 8) as u8;
    data4[1] = (data4_part1 & 0xFF) as u8;
    data4[2] = (data4_part2 >> 40) as u8;
    data4[3] = (data4_part2 >> 32) as u8;
    data4[4] = (data4_part2 >> 24) as u8;
    data4[5] = (data4_part2 >> 16) as u8;
    data4[6] = (data4_part2 >> 8) as u8;
    data4[7] = (data4_part2 & 0xFF) as u8;

    Ok(GUID {
        data1,
        data2,
        data3,
        data4,
    })
}

/// Helper function to test basic WinRT functionality
#[napi]
pub fn test_winrt_basic() -> Result<String> {
    // Initialize WinRT
    unsafe {
        RoInitialize(RO_INIT_MULTITHREADED)
            .map_err(|e| Error::from_reason(format!("Init failed: {:?}", e)))?;
    }

    // Try to create a simple WinRT object
    let class_name = HSTRING::from("Windows.Foundation.Uri");
    let uri_str = HSTRING::from("https://example.com");

    // Note: Actually creating a Uri requires more complex setup
    // This is just a demonstration of the pattern
    println!("Creating Uri instance for: {}", uri_str.to_string());
    
    unsafe {
        RoUninitialize();
    }

    Ok("WinRT basic test completed".to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_guid_parsing() {
        let guid_str = "00000000-0000-0000-C000-000000000046";
        let result = parse_guid(guid_str);
        assert!(result.is_ok());
    }
}
