#[cfg(test)]
mod tests {
    use windows::core;
    use windows::Win32::System::Com;

    #[test]
    fn test_co_create_instance() {
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
                            // obj.fmt()
                            let hstring = core::HSTRING::from("version");
                            let _rgsznames = core::PCWSTR::from_raw(hstring.as_ptr());
                            let mut _rgdispid = 0;
                            // let res = self.disp.GetIDsOfNames(
                            //     &self.iid,
                            //     &rgsznames,
                            //     1,
                            //     LOCALE_USER_DEFAULT,
                            //     &mut rgdispid,
                            // );
                            //    let res = obj.GetIDsOfNames(clsid,  1,0x1033, hstring,&mut rgdispid);
                            // let res = obj.GetIDsOfNames(clsid, 1, hstring, rgsznames, &mut rgdispid);
                            // 再获取数据
                           let res_typeinfo= obj.GetTypeInfo(0, 0x1033);
                           match res_typeinfo {
                               Ok(typeinfo)=>{
                                    println!("{:?}",typeinfo);
                                    // typeinfo.GetDllEntry(memid, invkind, pbstrdllname, pbstrname, pwordinal)
                                    // typeinfo.GetDocumentation(memid, pbstrname, pbstrdocstring, pdwhelpcontext, pbstrhelpfile)
                               }
                               Err(e)=>{
                                println!("{:?}",e)
                               }
                           }
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