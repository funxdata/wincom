#[cfg(test)]
mod tests {
    use windows::Win32::System::Com;
    use windows::core;

    #[test]
    fn test_cocreate_instance() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            // let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            // 根据appname 获取clsid
            let lpsz = core::HSTRING::from("powerpoint.application.16");
            let res = Com::CLSIDFromString(&lpsz);
            match res {
                Ok(clsid)=>{
                    println!("{:?}",clsid);
                    let res_obj:Result<Com::IDispatch, core::Error> = Com::CoCreateInstance(&clsid, None, Com::CLSCTX_SERVER);
                    match res_obj {
                        Ok(obj)=>{
                            println!("{:?}",obj);

                            // let aa = obj.as_unknown();
                            // aa.query(iid, interface)
                            // obj.query(iid, interface)
                            // println!()
                            // 获取
                            // let count = obj.GetTypeInfoCount();
                            // match count {
                            //     Ok(c)=>{
                            //         println!("count:{:?}",c);
                            //     }
                            //     Err(e)=>{
                            //         println!("err:{:?}",e);
                            //     }
                            // }

                            // // 先获取对象
                            let hstring = core::HSTRING::from("version");
                            let rgsznames = core::PCWSTR::from_raw(hstring.as_ptr());
                            let mut rgdispid = 0;
                            let res = obj.GetIDsOfNames(
                                &clsid,
                                &rgsznames,
                                1,
                                0x1033,
                                &mut rgdispid,
                            );
                            match res {
                                Ok(_)=>{
                                    println!("..........")
                                }
                                Err(e)=>{
                                    
                                    println!("GetIDsOfNames error:{:?}",e);
                                }
                            }

                            // 再获取数据
                            // obj.GetTypeInfo(itinfo, lcid)


                        }
                        Err(e)=>{
                            println!("{:?}",e);
                        }
                    }
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