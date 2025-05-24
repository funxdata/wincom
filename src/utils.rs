use std::mem::ManuallyDrop;
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::{VariantChangeType, VariantClear, VARIANT, VARIANT_0_0, VAR_CHANGE_FLAGS, VT_ARRAY, VT_BOOL, VT_BSTR, VT_EMPTY, VT_I2, VT_I4, VT_I8, VT_NULL, VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_UI8};
use windows::Win32::System::Variant;
use windows::Win32::System::Ole;

use windows::core;
pub trait VariantExt {
    fn null() -> VARIANT;
    fn from_i16(n: i16) -> VARIANT;
    fn from_i32(n: i32) -> VARIANT;
    fn from_i64(n: i64) -> VARIANT;
    fn from_u8(n: u8) -> VARIANT;
    fn from_u16(n: u16) -> VARIANT;
    fn from_vt4_u16(n: i32) -> VARIANT;
    fn from_u32(n: u32) -> VARIANT;
    fn from_u64(n: u64) -> VARIANT;
    fn from_f64(n: f64) -> VARIANT;
    fn from_vec_f64(arr: Vec<f64>) -> VARIANT;
    fn from_vec_default() -> VARIANT;
    fn from_vec_i64(arr: Vec<i64>) -> VARIANT;
    fn from_str<S: AsRef<str>>(s: S) -> VARIANT;
    fn from_bool(b: bool) -> VARIANT;
    fn to_i16(&self) -> core::Result<i16>;
    fn to_i32(&self) -> core::Result<i32>;
    fn to_i64(&self) -> core::Result<i64>;
    fn to_u8(&self) -> core::Result<u8>;
    fn to_u16(&self) -> core::Result<u16>;
    fn to_u32(&self) -> core::Result<u32>;
    fn to_u64(&self) -> core::Result<u64>;
    fn to_f64(&self) -> core::Result<f64>;
    fn to_string(&self) -> core::Result<String>;
    fn to_bool(&self) -> core::Result<bool>;
    fn to_f64_array(&self)-> core::Result<Vec<f64>>;
    fn to_double_array(&self)-> core::Result<Vec<f64>>;
    fn to_i64_array(&self)-> core::Result<Vec<i64>>;
    fn to_string_array(&self)-> core::Result<Vec<String>>;
    fn to_vart_array(&self)-> core::Result<Vec<VARIANT>>;
    fn to_idispatch(&self) -> core::Result<&IDispatch>;
    fn is_null(&self)->bool;
}

impl VariantExt for VARIANT {
    fn null() -> VARIANT {
        let mut variant = VARIANT::default();
        let v00 = VARIANT_0_0 {
            vt: VT_NULL,
            ..Default::default()
        };
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i16(n: i16) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I2,
            ..Default::default()
        };
        v00.Anonymous.iVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i32(n: i32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I4,
            ..Default::default()
        };
        v00.Anonymous.lVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i64(n: i64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I8,
            ..Default::default()
        };
        v00.Anonymous.llVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_u8(n: u8) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI1,
            ..Default::default()
        };
        v00.Anonymous.bVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_u16(n: u16) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI2,
            ..Default::default()
        };
        v00.Anonymous.uiVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_vt4_u16(n: i32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I4,
            ..Default::default()
        };
        v00.Anonymous.lVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }

    fn from_u32(n: u32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI4,
            ..Default::default()
        };
        v00.Anonymous.ulVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_u64(n: u64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_UI8,
            ..Default::default()
        };
        v00.Anonymous.ullVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_f64(n: f64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_R8,
            ..Default::default()
        };
        v00.Anonymous.dblVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_vec_f64(arr: Vec<f64>) -> VARIANT{
        unsafe {
            let mut variant = VARIANT::default();
            let mut v00 = VARIANT_0_0 {
                vt: Variant::VARENUM(VT_ARRAY.0|VT_R8.0),
                ..Default::default()
            };
            let psa = Ole::SafeArrayCreateVector(VT_R8, 0, arr.len() as u32);
            for (i, &value) in arr.iter().enumerate() {
                Ole::SafeArrayPutElement(psa, &(i as i32), &value as *const _ as *const _).ok().unwrap();
            }
            v00.Anonymous.parray = psa;
            variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
            variant
        }
    }
    fn from_vec_default() -> VARIANT{
        unsafe {
            let mut variant = VARIANT::default();
            let mut v00 = VARIANT_0_0 {
                vt: Variant::VARENUM(VT_ARRAY.0|VT_R8.0),
                ..Default::default()
            };
            let psa = Ole::SafeArrayCreateVector(VT_R8, 0, 2);
            // for (i, &value) in arr.iter().enumerate() {
            //     Ole::SafeArrayPutElement(psa, &(i as i32), &value as *const _ as *const _).ok().unwrap();
            // }
            v00.Anonymous.parray = psa;
            variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
            variant
        }
    }
    fn from_vec_i64(arr: Vec<i64>) -> VARIANT{
        unsafe {
            let mut variant = VARIANT::default();
            
            let mut v00 = VARIANT_0_0 {
                vt: VT_ARRAY,
                ..Default::default()
            };
            let psa = Ole::SafeArrayCreateVector(VT_I8, 0, arr.len() as u32);
            for (i, &value) in arr.iter().enumerate() {
                Ole::SafeArrayPutElement(psa, &(i as i32), &value as *const _ as *const _).ok().unwrap();
            }
            v00.Anonymous.parray = psa;
            variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
            variant
        }
    }
    fn from_str<S: AsRef<str>>(s: S) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BSTR,
            ..Default::default()
        };
        let bstr = core::BSTR::from(s.as_ref());
        v00.Anonymous.bstrVal = ManuallyDrop::new(bstr);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_bool(b: bool) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BOOL,
            ..Default::default()
        };
        v00.Anonymous.boolVal = VARIANT_BOOL::from(b);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn to_i16(&self) -> core::Result<i16> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I2)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.iVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_i32(&self) -> core::Result<i32> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.lVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_i64(&self) -> core::Result<i64> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_I8)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.llVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u8(&self) -> core::Result<u8> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI1)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.bVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u16(&self) -> core::Result<u16> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI2)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.uiVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u32(&self) -> core::Result<u32> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.ulVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_u64(&self) -> core::Result<u64> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_UI8)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.ullVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_f64(&self) -> core::Result<f64> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_R8)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.dblVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_string(&self) -> core::Result<String> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_BSTR)?;
            let v00 = &new.Anonymous.Anonymous;
            let str = v00.Anonymous.bstrVal.to_string();
            VariantClear(&mut new)?;
            Ok(str)
        }
    }
    fn to_bool(&self) -> core::Result<bool> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, VAR_CHANGE_FLAGS(0), VT_BOOL)?;
            let v00 = &new.Anonymous.Anonymous;
            let b = v00.Anonymous.boolVal.as_bool();
            VariantClear(&mut new)?;
            Ok(b)
        }
    }
    fn to_i64_array(&self)-> core::Result<Vec<i64>> {
        unsafe {
            let psa = self.Anonymous.Anonymous.Anonymous.parray;
            let count =  Variant::VariantGetElementCount(self) as i32;
            let mut array_vec = Vec::with_capacity(count as usize);
            for i in 0..count {
                let mut value: i64 = 0;
                Ole::SafeArrayGetElement(psa, &i, &mut value as *mut _ as *mut _)?;
                array_vec.push(value);
            }
            // Ok(vec)
            Ok(array_vec)
        }
    }
    fn to_f64_array(&self)-> core::Result<Vec<f64>> {
        unsafe {
            // println!("Unknown type: {:?}", self.Anonymous.Anonymous.vt);
            let psa = self.Anonymous.Anonymous.Anonymous.parray;
            let count =  Variant::VariantGetElementCount(self) as i32;
            let mut array_vec = Vec::with_capacity(count as usize);
            for i in 0..count {
                let mut value: f64 = 0.0;
                Ole::SafeArrayGetElement(psa, &i, &mut value as *mut _ as *mut _)?;
                array_vec.push(value);
            }
            // Ok(vec)
            Ok(array_vec)
        }
    }
    fn to_double_array(&self)-> core::Result<Vec<f64>> {
        unsafe {
            // println!("Unknown type: {:?}", self.Anonymous.Anonymous.vt);
            let psa = self.Anonymous.Anonymous.Anonymous.parray;
            let count =  Variant::VariantGetElementCount(self) as i32;
            let mut array_vec = Vec::with_capacity(count as usize);
            for i in 0..count {
                let mut value: f64 = 0.0;
                Ole::SafeArrayGetElement(psa, &i, &mut value as *mut _ as *mut _)?;
                array_vec.push(value);
            }
            // Ok(vec)
            Ok(array_vec)
        }
    }
    
    fn to_string_array(&self)-> core::Result<Vec<String>> {
        unsafe {
            let psa = self.Anonymous.Anonymous.Anonymous.parray;
            let count =  Variant::VariantGetElementCount(self) as i32;
            let mut array_vec = Vec::with_capacity(count as usize);
            for i in 0..count {
                let mut value: String = "".to_string();
                Ole::SafeArrayGetElement(psa, &i, &mut value as *mut _ as *mut _)?;
                array_vec.push(value);
            }
            // Ok(vec)
            Ok(array_vec)
        }
    }
    fn to_vart_array(&self)-> core::Result<Vec<VARIANT>> {
        unsafe {
            let _psa = self.Anonymous.Anonymous.Anonymous.parray;
            let count =  Variant::VariantGetElementCount(self) as i32;
            println!("。。。。。。。。。。。。。。。。。。::{:?}",count);
            println!("count::{:?}",count);
            println!("。。。。。。。。。。。。。。。。。。::{:?}",count);
            // let mut array_vec = Vec::with_capacity(count as usize);
            // for i in 0..count {
            //     let mut value: f64 = 0.0;
            //     Ole::SafeArrayGetElement(psa, &i, &mut value as *mut _ as *mut _)?;
            //     array_vec.push(value);
            // }
            // // Ok(vec)
            // Ok(array_vec)
            Ok(Vec::new())
        }


    }
    fn to_idispatch(&self) -> core::Result<&IDispatch> {
        unsafe {
            let v00 = &self.Anonymous.Anonymous;
            v00.Anonymous.pdispVal.as_ref().ok_or(core::Error::new(
                core::HRESULT(0x00123456),
                core::HSTRING::from("com-shim: Cannot read IDispatch"),
            ))
        }
    }
    fn is_null(&self)->bool{
        unsafe {
            if self.Anonymous.Anonymous.vt == VT_EMPTY ||  self.Anonymous.Anonymous.vt ==VT_NULL {
                return true;
            }
            return false;
        }
    }
}