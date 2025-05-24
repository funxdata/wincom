#![allow(non_camel_case_types,non_snake_case, unused,non_upper_case_globals)] 
#[cfg(test)]
mod tests {
    use dyteslogs::logs::LogError;
    use windows::Win32::System::{Com, Variant};
    use windows::Win32::System::Variant::{VARIANT, VARIANT_0_0, VT_ARRAY, VT_BOOL, VT_I4, VT_R8, VT_SAFEARRAY, VT_UI8};
    use windows::Win32::System::Ole;
    use wincom::{dispatch::ComObject, utils::VariantExt};
    use std::mem::ManuallyDrop;

    #[test]
    fn test_win_variant_array() {
 // Initialize COM library
        unsafe { 
            Com::CoInitialize(None).unwrap();
            println!("..............");
               // Create a SafeArray of integers (VT_I4) with 10 elements
            // let test_arr = Ole::SafeArrayCreateVector(VT_R8, 1, 10);
            
            // let hr = Ole::SafeArrayGetVartype(test_arr);
            // println!("type:{:?}",hr);
            // // let count =  Variant::VariantGetElementCount(test_arr) as i32;
            // // println!("{:?}",count);
            // let hr = Ole::SafeArrayGetDim(test_arr);
            // println!("type:{:?}",hr);
            // let mut vart: VARIANT = VARIANT::default();
            // Variant::VariantChangeType(&mut vart, self, VAR_CHANGE_FLAGS(0), VT_R8).unwrap();
            // let v00 = &new.Anonymous.Anonymous;
            // let b = v00.Anonymous.boolVal.as_bool();
            let mut variant = VARIANT::default();
            let mut v00 = VARIANT_0_0 {
                vt: VT_R8,
                ..Default::default()
            };
            v00.Anonymous.dblVal = 3.14;
            variant.Anonymous.Anonymous = ManuallyDrop::new(v00);

            let vt = variant.Anonymous.Anonymous.vt;
            let type_name = match vt {
                Variant::VT_EMPTY => "VT_EMPTY",
                Variant::VT_NULL => "VT_NULL",
                Variant::VT_I2 => "VT_I2",
                Variant::VT_I4 => "VT_I4",
                Variant::VT_R4 => "VT_R4",
                Variant::VT_R8 => "VT_R8",
                Variant::VT_CY => "VT_CY",
                Variant::VT_DATE => "VT_DATE",
                Variant::VT_BSTR => "VT_BSTR",
                Variant::VT_DISPATCH => "VT_DISPATCH",
                Variant::VT_ERROR => "VT_ERROR",
                Variant::VT_BOOL => "VT_BOOL",
                Variant::VT_VARIANT => "VT_VARIANT",
                Variant::VT_UNKNOWN => "VT_UNKNOWN",
                Variant::VT_DECIMAL => "VT_DECIMAL",
                Variant::VT_I1 => "VT_I1",
                Variant::VT_UI1 => "VT_UI1",
                Variant::VT_UI2 => "VT_UI2",
                Variant::VT_UI4 => "VT_UI4",
                Variant::VT_I8 => "VT_I8",
                Variant::VT_UI8 => "VT_UI8",
                Variant::VT_INT => "VT_INT",
                Variant::VT_UINT => "VT_UINT",
                Variant::VT_VOID => "VT_VOID",
                Variant::VT_HRESULT => "VT_HRESULT",
                Variant::VT_PTR => "VT_PTR",
                Variant::VT_SAFEARRAY => "VT_SAFEARRAY",
                Variant::VT_CARRAY => "VT_CARRAY",
                Variant::VT_USERDEFINED => "VT_USERDEFINED",
                Variant::VT_LPSTR => "VT_LPSTR",
                Variant::VT_LPWSTR => "VT_LPWSTR",
                Variant::VT_RECORD => "VT_RECORD",
                Variant::VT_INT_PTR => "VT_INT_PTR",
                Variant::VT_UINT_PTR => "VT_UINT_PTR",
                Variant::VT_FILETIME => "VT_FILETIME",
                Variant::VT_BLOB => "VT_BLOB",
                Variant::VT_STREAM => "VT_STREAM",
                Variant::VT_STORAGE => "VT_STORAGE",
                Variant::VT_STREAMED_OBJECT => "VT_STREAMED_OBJECT",
                Variant::VT_STORED_OBJECT => "VT_STORED_OBJECT",
                Variant::VT_BLOB_OBJECT => "VT_BLOB_OBJECT",
                Variant::VT_CF => "VT_CF",
                Variant::VT_CLSID => "VT_CLSID",
                Variant::
                VT_VERSIONED_STREAM => "VT_VERSIONED_STREAM",
                _ => "Unknown",
            };

            println!("The variant is of type: {}", type_name);

            // 设置 VARIANT 的类型和值（这里设置为 VT_R8 作为示例）
            // vart.Anonymous.Anonymous.vt = windows::Win32::System::Variant::VARENUM(VT_R8.0 as u16);
            // vart.Anonymous.Anonymous.Anonymous.dblVal = 3.14;


            Com::CoUninitialize();
        }

    }
}