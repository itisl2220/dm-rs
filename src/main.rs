use dm_rs::DmResult;
use dm_rs::DmSoft;

fn main() -> DmResult<()> {
    let obj = DmSoft::new();
    match obj {
        Ok(obj) => {
            let ver = obj.ver()?;
            println!("ver:{:?}", ver);
            let v = obj.reg("mh84909b3bf80d45c618136887775ccc90d27d7", "mmqnvy80ddz0ec7")?;
            println!("v:{:?}", v);
            let res = obj.enable_pic_cache()?;
            println!("res:{:?}", res);
        }
        Err(e) => {
            println!("{:?}", e);
        }
    }
    Ok(())
}
