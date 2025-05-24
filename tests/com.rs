#[cfg(test)]
mod tests {
    use windows::{core, Win32::System::Com};

    #[test]
    fn test_clsidfromstring() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            // let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            // 根据appname 获取clsid
            let lpsz = core::HSTRING::from("powerpoint.application");
            let res = Com::CLSIDFromString(&lpsz);
            match res {
                Ok(clsid)=>{
                    println!("{:?}",clsid);
                }
                Err(e)=>{
                    println!("{:?}",e);
                }
            }
            Com::CoUninitialize();
      }
    }
    
    #[test]
    fn test_clsidfrom_prog_id() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            // let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            // 根据appname 获取clsid
            let lpsz = core::HSTRING::from("powerpoint.application");
            let res = Com::CLSIDFromProgID(&lpsz);
            match res {
                Ok(clsid)=>{
                    println!("{:?}",clsid);
                }
                Err(e)=>{
                    println!("{:?}",e);
                }
            }
            Com::CoUninitialize();
      }
    }

}