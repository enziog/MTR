extern crate calamine;
extern crate glob;

use std::io;
use std::process;
//use std::path::Path;
use glob::glob;
//use std::vec;
use calamine::{Reader, Xlsx, open_workbook};
//use calamine::{Reader, open_workbook_auto, Xlsx, DataType, open_workbook};

fn run() {
    let mut mcl_path_list = Vec::new();
    mcl_path_list = find_mcl();
    println!( "Received MCL path list is:\n{:?}\n是否继续?点击回车继续.输入\"n\"退出重新查找", mcl_path_list);
    let tnl: Vec<String> = generate_tnl(mcl_path_list);
    find_mtr(&tnl);
}
fn generate_tnl(mcl_path_list: Vec<String>) -> Vec<String>{
    let mut tnl: Vec<String> = Vec::new();
    for mcl_path in mcl_path_list {
        let mut excel: Xlsx<_> = open_workbook(&mcl_path).unwrap();
        if let Some(Ok(r)) = excel.worksheet_range("Recovered_Sheet1") {
            for row in r.rows() {
                let row = format!("{:?}", row[3]);
                //println!("row[3]={:?}", row);
                if row != "Empty" {
                    if row.find("\n") == None {
                        //println!("found a cell with more than 1 tn {:?}", row);
                        let mut cl = Vec::<&str>::new();
                        cl = row.split("\"").collect();
                        for e in cl {   //分割条件
                            let f = e.find("String") != None || e.find(")") != None || e.find("Coded") != None;
                            //println!("f is {}", f);
                            match f {
                                true => continue,
                                false => {
                                            if e.find("\\r\\n") != None {
                                                println!("{}", e);
                                                let mut ml : Vec<&str> = e.split("\\r\\n").collect();
                                                for elem in ml {
                                                    //println!("{}", elem);
                                                    tnl.push(elem.to_string());
                                                }
                                            } else {
                                                tnl.push(e.to_string());
                                            }
                                         }
                            }
                        }
                    }
                } else {
                    continue;
                }
            }

        }
             println!("{:?}", tnl);
    }
    tnl
}

fn find_mtr(tnl: &Vec<String>) {
    //collect MTRs based on tnlist
    for tn in tnl {
        let mut glob_mtr = String::from("*") + &tn + "*";
        println!("glob_mtr is: {}", glob_mtr);
    }
}
fn find_mcl() -> Vec<String> {
    let mut job_no = String::new();
    io::stdin().read_line(&mut job_no)
        .expect("Failed to read job_no");
    let job_no = job_no.trim();
    let exit = String::from("q");
    if &job_no == &exit {
        println!("exit code detected\nI am going to take a rest...");
        process::exit(0);
    } else {
        println!("Job No: {}\nI am trying to find tns listed in MCL...", job_no);
    }
    if &job_no == &"" {
        println!("Seems you are tricking me, please do not left job No. blank.\nLet's do it again");
        run();
    }

    let mut mcl_path = String::from("/home/enziog/Rust/FindMTR/**/");
    mcl_path = mcl_path + "*" + job_no + "*.xlsx";
    let mut mcl_path_list:Vec<String> = Vec::new();
//    println!("Length of MCL Path List is: {}", mcl_path_list.len()); 
//考虑改进glob查找条件以便能在输入不完整工作令时还能查找到相关的Excel: 待测试，如果条件输入有问题，找到太多Excel时应暂停程序，不要再生成tn
//mcl_path_list直接用Vec<Path>，省去转换？
    for entry in glob(&mcl_path).unwrap() {
        match entry {
            Ok(path) => {mcl_path_list.push(path.to_str().unwrap().to_owned()); },
            //println!("{:?}", path.display()); 
            Err(e)   => {println!("{:?}", e); run();},
        }
    }

    println!( "{:?}", mcl_path_list);
//    println!("Length of MCL Path List is: {}", mcl_path_list.len()); 

    if mcl_path_list.len() == 0 {
        println!("But I can't find xlsx document of specified Job No");
    }
    mcl_path_list
}
fn main() {
    println!("Hello, I am FindMTR program. Please input the job_no and I will collect the MTR for you based on tracking Nos in MCL\nIf you want to quit, please press \"q\" on keyboard and then press \"ENTER\"");
    run();
    process::exit(0);
}

