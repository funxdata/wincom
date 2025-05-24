#[cfg(test)]
mod tests {
    use dyteslogs::logs::LogError;
    use windows::Win32::System::Com;
    use wincom::{dispatch::ComObject, utils::VariantExt};
    use windows::core::GUID;

    #[test]
    fn test_app_autocad() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            // let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            // 根据appname 获取clsid
            // let lpsz = core::HSTRING::from("autocad.application");
           let app =  ComObject::new_from_name("autocad.application", GUID::zeroed()).unwrap();
           let ver = app.get_property("version").log_error("got property error").unwrap();
           let ver_str = ver.to_string().unwrap();
           println!("--{:?}--",ver_str);
           let app =  ComObject::new_from_name("autocad.application", GUID::zeroed()).unwrap();
           let width = app.get_property("width").log_error("got property error").unwrap();
           let width_str = width.to_u32().unwrap();
           println!("--{:?}--",width_str);
           // 执行退出
           let args = Vec::new();
           let res_data = app.invoke_method("", args);
           match res_data {
               Ok(data)=>{
                    println!("{:?}",data.to_string().unwrap());
               }
               Err(e)=>{
                    println!("{:?}",e);
               }
            }

           // Com::CoCreateInstance(&rclsid, None, Com::CLSCTX_ALL)
           Com::CoUninitialize();
      }
    }

}