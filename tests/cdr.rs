#![allow(non_camel_case_types,non_snake_case, unused,non_upper_case_globals)] 
#[cfg(test)]
mod tests {
    use wincom::dispatch::ComObject;
    use wincom::utils::VariantExt;
    use windows::Win32::System::Com;
    use windows::core::GUID;
    use windows::core;
    use windows::Win32::System::Ole;

    #[test]
    fn test_app_cdr() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
           let lpsz = core::HSTRING::from("coreldraw.application");
           let rclsid =   Com::CLSIDFromProgID(&lpsz).unwrap();
           println!("{:?}",rclsid);
        //    ole:GetActiveObject(rclsid, pvreserved, ppunk)
           let mut app_opt:Option<core::IUnknown> = None; 
           let hr = Ole::GetActiveObject(&rclsid,None,&mut app_opt);
           println!("{:?}",app_opt);
           Com::CoUninitialize();
      }
    }
    #[test]
    fn test_app_cdr_docs() {
        unsafe{
           Com::CoInitialize(None).expect("init com error");
           let app =  ComObject::new_from_name("autocad.application", GUID::zeroed()).unwrap();
           let height_var = app.get_property("height").unwrap();
           let height =  height_var.to_u32().unwrap();
           println!("{:?}",height);
           Com::CoUninitialize();
      }
    }
}