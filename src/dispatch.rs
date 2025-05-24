#![allow(dead_code)]
use std::mem::ManuallyDrop;
use windows::core::ComInterface;
use windows::core::IUnknown;
use windows::core::GUID;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant;
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::System::Variant::VARIANT_0_0;
use windows::Win32::System::Variant::VT_BYREF;
use windows::Win32::System::Variant::VT_DISPATCH;
use windows::core;
use windows::Win32::System::Com;
use windows::Win32::System::Ole;
use windows::core::PCWSTR;
use windows::Win32::Globalization::GetSystemDefaultLCID;
use windows::Win32::System::Variant::VT_VARIANT;

use super::utils::VariantExt;
use dyteslogs::logs::LogError;

pub struct ComObject {
    iid: core::GUID,
    disp: Com::IDispatch,
}

#[allow(unused)]
impl ComObject {    
    // 创建
    pub fn new_from_name(id: &str,iid:core::GUID) -> Option<Self> {
        unsafe {
            let lpsz = core::HSTRING::from(id);
            let rclsid_res = Com::CLSIDFromProgID(&lpsz);
            if rclsid_res.is_err(){
                return None;
            }
            let rclsid = rclsid_res.unwrap();
            let ikun_server:Result<IUnknown, core::Error> =  Com::CoCreateInstance(&rclsid, None, Com::CLSCTX_LOCAL_SERVER);
            if ikun_server.is_err() {
                return None;
            }
            let ikun = ikun_server.unwrap();
            let idisp =  ikun.cast().unwrap();
            return Some(Self { 
                iid:rclsid,
                disp:idisp 
            });
        }
    }
    // 创建
    pub fn new_from_app(id: &str) -> Option<Self> {
        unsafe {
            let lpsz = core::HSTRING::from(id);
            if let Some(rclsid) = Com::CLSIDFromProgID(&lpsz).log_error("got progclsidid error"){
                let mut app_opt:Option<core::IUnknown> = None; 
                let hr = Ole::GetActiveObject(&rclsid,None,&mut app_opt);
                if hr.is_err() {
                    return None;
                }
                let app_ikun = app_opt.unwrap();
                let app_idispatch:IDispatch = app_ikun.cast().unwrap();
                return Some(Self { 
                    iid:rclsid,
                    disp:app_idispatch 
                });
            }
            None
        }
    }
    
    // 启动目录创建引用
    pub fn new_from_path(id: &str) -> Option<Self> {
        unsafe {
            let lpsz = core::HSTRING::from(id);
            if let Some(rclsid) = Com::CLSIDFromProgID(&lpsz).log_error("got progclsidid error"){
                let mut app_opt:Option<core::IUnknown> = None; 
                let hr = Ole::GetActiveObject(&rclsid,None,&mut app_opt);
                if hr.is_err() {
                    return None;
                }
                let app_ikun = app_opt.unwrap();
                let app_idispatch:IDispatch = app_ikun.cast().unwrap();
                return Some(Self { 
                    iid:rclsid,
                    disp:app_idispatch 
                });
            }
            None
        }
    }



    pub fn clone_from(data:Com::IDispatch,iid:core::GUID)->core::Result<Self> {
        unsafe {
            Ok(Self { 
                iid:iid,
                disp:data 
            })
        }
    }

    fn invoke(
        &self,
        dispidmember: i32,
        pdispparams: &Com::DISPPARAMS,
        wflags: Com::DISPATCH_FLAGS,
    ) -> core::Result<VARIANT> {
        unsafe {
            let mut result = VARIANT::null();
            let res = self.disp.Invoke(
                dispidmember,
                &core::GUID::zeroed(),
                0x0,
                wflags,
                pdispparams,
                Some(&mut result),
                None,
                None,
            );
            match res {
                Ok(res_data)=>{
                    return Ok(result);
                }
                Err(e)=>{
                    return Err(e);
                }
            }
        }
    }

    pub fn invoke_callback(&self,method:&str,mut args:Vec<VARIANT>)->core::Result<Vec<VARIANT>>{
        let mut params:[VARIANT;2] =[VARIANT::default(),VARIANT::default()];
        let mut min_p: VARIANT = VARIANT::default();
        let mut max_p: VARIANT = VARIANT::default();
        // params[0].Anonymous.Anonymous.Anonymous.parray = Ole::SafeArrayCreateVector(VT_R8, 0, arr.len() as u32);
        unsafe {
            (*params[0].Anonymous.Anonymous).vt = Variant::VARENUM(VT_BYREF.0| VT_VARIANT.0);
            (*params[0].Anonymous.Anonymous).Anonymous.byref = &mut min_p as *mut _ as *mut _;                
            (*params[1].Anonymous.Anonymous).vt = Variant::VARENUM(VT_BYREF.0| VT_VARIANT.0);
            (*params[1].Anonymous.Anonymous).Anonymous.byref = &mut max_p as *mut _ as *mut _;                   
        }
        let dispidmember = self.get_id_from_name(method).log_error("got vart err").unwrap();
        let mut dp = Com::DISPPARAMS::default();
        dp.cArgs =params.len() as u32;
        dp.rgvarg =params.as_mut_ptr();
        let mut result = VARIANT::null();
        let res = unsafe { 
            self.disp.Invoke(
            dispidmember,
            &core::GUID::zeroed(),
            0x0,
            Com::DISPATCH_METHOD,
            &mut dp,
            Some(&mut result),
            None,
            None,
            )
        };
        match res {
            Ok(res_data)=>{
                // return Ok(result);
               args[0] = min_p.clone();
               args[1] = max_p.clone();
                return Ok(args);
            }
            Err(e)=>{
                return Err(e);
            }
        }
    }
    
    fn invoke_array(
        &self,
        dispidmember: i32,
        pdispparams: &mut Com::DISPPARAMS,
        wflags: Com::DISPATCH_FLAGS,
    ) -> core::Result<VARIANT> {
        unsafe {
            let mut result = VARIANT::null();
            let res = self.disp.Invoke(
                dispidmember,
                &core::GUID::zeroed(),
                0x0,
                wflags,
                pdispparams,
                Some(&mut result),
                None,
                None,
            );
            match res {
                Ok(res_data)=>{
                    return Ok(result);
                }
                Err(e)=>{
                    return Err(e);
                }
            }
        }
    }

    fn get_id_from_name(&self,name:&str)->core::Result<i32> {
        unsafe{
            let lcid = GetSystemDefaultLCID();
            let hstring = core::HSTRING::from(name);
            let rgsz_names = PCWSTR::from_raw(hstring.as_ptr());
            let mut rgdispid = 0;
            let res = self.disp.GetIDsOfNames(
                &GUID::zeroed(),
                &rgsz_names,
                1,
                lcid,
                &mut rgdispid,
            );
            Ok(rgdispid)
        }
    }
    

    // 获取参数
    pub fn get_property(&self, prop: &str) -> core::Result<VARIANT> {
        let dispidmember = self.get_id_from_name(prop).log_error("got vart err").unwrap();
        let mut dp = Com::DISPPARAMS::default();
        let mut args = vec![];
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_mut_ptr();
        self.invoke(dispidmember, &dp, Com::DISPATCH_PROPERTYGET)
    }
    // 获取参数
    pub fn get_property_by_vart(&self, prop: &str,mut param:VARIANT) -> core::Result<VARIANT> {
        let dispidmember = self.get_id_from_name(prop).log_error("got vart err").unwrap();
        let mut dp = Com::DISPPARAMS::default();
        let mut args = vec![param];
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_mut_ptr();
        self.invoke(dispidmember, &dp, Com::DISPATCH_PROPERTYGET)
    }

    // 设置参数
    pub fn set_property(&self, prop: &str, mut args: Vec<VARIANT> ) -> core::Result<()> {
        let dispidmember = self.get_id_from_name(prop).log_error("got vart err").unwrap();
        let mut dp = Com::DISPPARAMS::default();
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_mut_ptr();
        dp.cNamedArgs = 1;
        let mut named_args = vec![Ole::DISPID_PROPERTYPUT];
        dp.rgdispidNamedArgs = named_args.as_mut_ptr();
        self.invoke(dispidmember, &mut dp, Com::DISPATCH_PROPERTYPUT)?;
        Ok(())
    }

    // 执行回调方法
    pub fn invoke_method(&self, method: &str, mut args: Vec<VARIANT>) -> core::Result<VARIANT> {
        let dispidmember = self.get_id_from_name(method).log_error("got vart err").unwrap();
        let mut dp = Com::DISPPARAMS::default();
        args.reverse();
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_mut_ptr();
        self.invoke(dispidmember, &dp, Com::DISPATCH_METHOD)
    }

    
    // 执行回调方法
    pub fn callback_invoke_method(&self, method: &str, mut args: Vec<VARIANT>) -> core::Result<Vec<VARIANT>> {
        let dispidmember = self.get_id_from_name(method).log_error("got vart err").unwrap();
        let mut dp = Com::DISPPARAMS::default();
        args.reverse();
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_mut_ptr();
        let res = self.invoke_array(dispidmember, &mut dp, Com::DISPATCH_METHOD);
        let res_data = args.clone();
        Ok(res_data)
    }

    //  转换
    pub fn get_variant(&self) -> core::Result<VARIANT>{
            let mut vart = VARIANT::default();
            let mut v00 = VARIANT_0_0 {
                vt: VT_DISPATCH,
                ..Default::default()
            };
            v00.Anonymous.pdispVal = ManuallyDrop::new(Some(self.disp.clone()));
            vart.Anonymous.Anonymous = ManuallyDrop::new(v00);
            Ok(vart)
    }
}
