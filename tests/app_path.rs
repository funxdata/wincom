#[cfg(test)]
mod tests {
    use windows::Win32::System::Com;
    use windows::core;

    #[test]
    fn test_app_from_path() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            // let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            // 根据appname 获取clsid
            // let lpsz = core::HSTRING::from("autocad.application");
            println!("hello");
            // let obj = Ole::
            // let app_apth = 
            // let app_apth = core::HSTRING::from(r"C:\\Program Files\\Corel\\CorelDRAW Graphics Suite\\25\\Programs64\\CorelDRW.exe");
            // let app_res = Com::CoGetObject(&app_apth,None);
            // match app_res {
            //     Ok(app)
            // }

            let lpsz = core::HSTRING::from(r"C:\\Program Files\\Corel\\CorelDRAW Graphics Suite\\25\\Programs64\\CorelDRW.exe");
            let rclsid = Com::CLSIDFromProgID(&lpsz).unwrap();
            println!("{:?}",rclsid);

           // Com::CoCreateInstance(&rclsid, None, Com::CLSCTX_ALL)
           Com::CoUninitialize();
      }
    }

}