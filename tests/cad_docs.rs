#[cfg(test)]
mod tests {
    use dyteslogs::logs::LogError;
    use windows::Win32::System::Com;
    use wincom::{dispatch::ComObject, utils::VariantExt};
    use windows::core::GUID;

    #[test]
    fn test_app_autocad_docs() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            // let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            // 根据appname 获取clsid
            // let lpsz = core::HSTRING::from("autocad.application");
           let app =  ComObject::new_from_name("autocad.application", GUID::zeroed()).unwrap();
           let docs_var = app.get_property("documents").log_error("got property error").unwrap();
           let docs_dispatch = docs_var.to_idispatch().unwrap();
           let info_res = docs_dispatch.GetTypeInfo(0,0);
           match info_res{
                Ok(info)=>{
                    // info.
                     // Get type attributes
                    // let mut type_attr: *mut TYPEATTR = std::ptr::null_mut();
                    let hr_type_attr = info.GetTypeAttr().unwrap();
                    println!("{:?}",hr_type_attr);
                    let type_attr = *hr_type_attr;

                    // Display the GUID
                    println!("GUID: {:?}", type_attr.guid);
                    
                }
                Err(e)=>{
                    println!("{:?}",e);
                }
           }
        //    docs_dispatch.
           let docs = ComObject::clone_from(docs_dispatch.clone(), GUID::zeroed()).unwrap();
            let sum_var = docs.get_property("count").unwrap();

           let sum = sum_var.to_u32().unwrap();
           println!("{:?}",sum);


           // Com::CoCreateInstance(&rclsid, None, Com::CLSCTX_ALL)
           Com::CoUninitialize();
      }
    }

}