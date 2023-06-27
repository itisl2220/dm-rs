use dm_rs::DmSoft;

fn main() -> Result<(), anyhow::Error> {
    let obj = DmSoft::new();
    match obj {
        Ok(obj) => {
            let ver = obj.ver()?;
            println!("ver:{:?}", ver);
            let v = obj.reg("mh84909b3bf80d45c618136887775ccc90d27d7", "mmqnvy80ddz0ec7")?;
            println!("v:{:?}", v);
        }
        Err(e) => {
            println!("{:?}", e);
        }
    }
    Ok(())
}
