#[cfg(test)]
mod tests {
    use wincom::dispatch::ComObject;
    use windows::{core::GUID, Win32::System::Com};

    #[test]
    fn test_init() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            
            let res_data=  ComObject::new_from_name("coreldraw.application", app_inter_iid);
            match res_data {
                Some(_)=>{
                    // obj.get_property(prop)
                    // println!("{:?}",obj.);
                    // println!(".......{:?}",obj.);
                }
                None=>{
                    print!("初始化失败")
                }
            }     
            Com::CoUninitialize();
      }
        // println!()
    }
}