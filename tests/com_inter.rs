#[cfg(test)]
mod tests {
    use windows::{core::{self, ComInterface, Error, IUnknown, PCWSTR}, Win32::System::Com::{self, IDispatch}};
    use windows::core::GUID;

    #[test]
    fn test_co_create_instance() {
        unsafe{
            Com::CoInitialize(None).expect("init com error");
            // let app_inter_iid = GUID::from_u128(0xB058000B_9AA4_44FD_9547_4F91EB757AC4);
            // 根据appname 获取clsid
            let lpsz = core::HSTRING::from("autocad.application");
            let res = Com::CLSIDFromString(&lpsz);
            if res.is_err(){
                return
            }
            // println!("{:?}",res.unwrap());
            let guid = res.unwrap();
            // let app = ComObject::new_from_name("autocad.application", guid);
            let app_res:Result<IUnknown,Error> = Com::CoCreateInstance(&guid, None, Com::CLSCTX_LOCAL_SERVER);
            println!("{:?}",app_res);
            // if app_res.is_err() {
            // }
            let cominterface = app_res.unwrap();
            // cominterface.query(iid, interface)
            // cominterface.cast()
            // let mut inter:IDispatch;
            // let hr = cominterface.query(&guid, &mut inter);
            // 查询 IDispatch 接口
            let cad_disp: IDispatch = cominterface.cast().unwrap();
            // 调用 GetIDsOfNames
            let name = "version".to_string();
            let hstring = core::HSTRING::from(name);
            let rgsznames = PCWSTR::from_raw(hstring.as_ptr());

            let mut rgdispid = 0;
            cad_disp.GetIDsOfNames(&GUID::zeroed(), &rgsznames, 1, 0x400, &mut rgdispid).unwrap();
            println!("{:?}",rgdispid);

            // 设置 DISPPARAMS 结构
            // let mut params = DISPPARAMS {
            //     rgvarg: std::ptr::null_mut(),
            //     rgdispidNamedArgs: std::ptr::null_mut(),
            //     cArgs: 0,
            //     cNamedArgs: 0,
            // };

            // // 设置返回值
            // let mut result: VARIANT = std::mem::zeroed();
            // VariantInit(&mut result);

            // // 调用 Invoke
            // dispatch.Invoke(
            //     dispid,
            //     &GUID::zeroed(),
            //     0,
            //     windows::Win32::System::Com::DISPATCH_METHOD,
            //     &mut params,
            //     &mut result,
            //     std::ptr::null_mut(),
            //     std::ptr::null_mut(),
            // )?;

            // 打印结果
            println!("Successfully invoked method.");

            // 使用 IDispatch 接口
            println!("Successfully queried IDispatch interface.");


            // Com::CoCreateInstance(&rclsid, None, Com::CLSCTX_ALL)
            Com::CoUninitialize();
      }
    }

}