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
    fn test_variant_f32_array() {
 // Initialize COM library
        unsafe { 
            Com::CoInitialize(None).unwrap();
            println!("..............");
            let arr:[f64;3] = [17.20,18.60,0.00];
            let vart_arr = <VARIANT as VariantExt>::from_vec_f64(arr.to_vec());
            let vart_get_arr = vart_arr.to_f64_array().unwrap();
            println!("array info :  {:?}",vart_get_arr);
            // let vt_arr = Variant::VECME VT_ARRAY as u32;
            // let aa = (VT_ARRAY | VT_R8 as u8 );
            // println!("sum:::{:?}",aa);
            // let arr_i32:Vec<i64> = vec![17,18];
            // let vart_arr_i32 = <VARIANT as VariantExt>::from_vec_i64(arr_i32);
            // let vart_get_arr_i32 = vart_arr_i32.to_i64_array().unwrap();
            // println!("array info: {:?}",vart_get_arr_i32);
            Com::CoUninitialize();
        }
    }

    #[test]
    fn test_variant_f64_array() {
 // Initialize COM library
        unsafe { 
            Com::CoInitialize(None).unwrap();
            let mut sa = Ole::SafeArrayCreateVector(VT_R8, 0, 5); // Create a safe array of 5 doubles

            let mut p_data: *mut f64 = std::ptr::null_mut();
            unsafe {
                Ole::SafeArrayAccessData(sa, &mut p_data as *mut _ as *mut _).unwrap();
                for i in 0..5 {
                    *p_data.add(i) = i as f64; // Fill the array with some values
                }
                Ole::SafeArrayUnaccessData(sa);
            }
        
            // let mut vart = VARIANT::default();
            // vart.n1.n2_mut().vt = VT_ARRAY | VT_R8 as u16;
            // vart.n1.n2_mut().parray = sa;
            Com::CoUninitialize();
        }
    }

}