const fs = require('fs')
const path = require('path')
const xlsx = require('node-xlsx')
const dragWrapper = document.querySelector("#file-drag-test")
var  showinfo = document.querySelector("#show-info")

// const deviceType = {
//     'Apache': /^APACHE_/,
//     'Tomcat': /^TOMCAT_/,
//     'Weblogic': /^WEBLOGIC_/,
//     'Nginx': /^NGINX_/,
//     'Tuxedo': /^TUXEDO_/,
//     'MQ': /^MQ_/,
//     'MSSQL': /^MSSQL_/,
//     'MySQL': /^MYSQL_/,
//     'Oracle': /^ORACLE_/,
//     'Redis': /^REDIS_/,
//     'Suse': /^OS_LINUX_/,
//     "Windows": /^OS_WINDOWS_/
// }
const deviceType = {
    "Windows": /^OS_WINDOWS_/,
    'Suse': /^OS_LINUX_/,
    'Oracle': /^ORACLE_/,
    'Redis': /^REDIS_/,
    'MySQL': /^MYSQL_/,
    'MSSQL': /^MSSQL_/,
    'Apache': /^APACHE_/,
    'Tuxedo': /^TUXEDO_/,
    'Weblogic': /^WEBLOGIC_/,
    'IBM-MQ': /^MQ_/,
    'Nginx': /^NGINX_/,
    'Tomcat': /^TOMCAT_/ 
}

dragWrapper.addEventListener('drop',(e)=>{
    e.preventDefault()
    const files = e.dataTransfer.files
    if( files && files.length >0 ){
        const path = files[0].path
        // console.log(path)
        window.onload
        var oldStr = showinfo.innerHTML != null ? showinfo.innerHTML : ""
        showinfo.innerHTML = oldStr + "<br>[Info] 解析Excel文件成功."+path
        var ret = readXlsx(path)
        console.log(ret.length)
        if(ret.length<1){
            showinfo.innerHTML = showinfo.innerHTML+ "<br>[Warning] 数据解释失败，未发现sheet"
            return
        }
        var data = ret[0]
        showinfo.innerHTML = showinfo.innerHTML + "<br>[Info] 发现sheet："+ data.name

        showinfo.innerHTML = showinfo.innerHTML + "<br>[Info] 解析数据条数："+ data.data.length+ " 条"
        console.log(data)
        if(data.data[0].length<6 && data.data[0][5]!=="检查结果"){
            showinfo.innerHTML = showinfo.innerHTML + "<br>[Error] Excel 格式与目标格式不符！"
            return
        }
        //数据分类并分析
        var dataformat = dataParse(data.data)
        //导出数据
        var xlsobj = changeXlsxObj(dataformat)
        report(xlsobj)
    }
})
dragWrapper.addEventListener('dragover',e=>{
    e.preventDefault()
})
// 读取excel
function readXlsx(path){
    var list = xlsx.parse(path)
    return list 
}
// 数据处理
/*
var dataformat = {
    "分行名称":{
        "devicetype":[0,0]
        "devicetype":[0,0]
    },
}


*/
// 获取类型
function getDeviceName(v){
    for(var k in deviceType){
        if(deviceType[k].test(v)){
            return k
        }
    }
    return ""
}
function dataParse(data){

    var dataformat = {}
    // console.log(dataformat.hasOwnProperty('ad')) 判断key是否存在
    const arrLength = data.length
    for(var i=0; i<arrLength;i++){
        if(i==0){
            continue
        }
        var key = data[i][0]
        // console.log(data[i][5]==='OK')
        if(!dataformat.hasOwnProperty(key)){
            dataformat[key]={}
        }
        var dtype = getDeviceName(data[i][4])
        if(!dataformat[key].hasOwnProperty(dtype)){
            dataformat[key][dtype]={'success':0,'failed':0,'except':0}
        }
        // console.log(data[i][5]+'11111')
        if(data[i][5] === 'OK'){
            // console.log(111)
            dataformat[key][dtype]['success']=dataformat[key][dtype]['success']+1
        }else if(data[i][5] === false){
            // console.log(2222)
            dataformat[key][dtype]['failed']=dataformat[key][dtype]['failed']+1
        }else{
            // console.log(333333)
            dataformat[key][dtype]['except']=dataformat[key][dtype]['except']+1
            showinfo.innerHTML =showinfo.innerHTML+"<br>[Error] "+key+" "+dtype+"处理失败.  "+data[i][2]
        }
    }
    // console.log(dataformat)
    return dataformat
}
/**
 *对Date的扩展，将 Date 转化为指定格式的String
 *月(M)、日(d)、小时(h)、分(m)、秒(s)、季度(q) 可以用 1-2 个占位符，
 *年(y)可以用 1-4 个占位符，毫秒(S)只能用 1 个占位符(是 1-3 位的数字)
 *例子：
 *(new Date()).Format("yyyy-MM-dd hh:mm:ss.S") ==> 2006-07-02 08:09:04.423
 *(new Date()).Format("yyyy-M-d h:m:s.S")      ==> 2006-7-2 8:9:4.18
 */
Date.prototype.format = function (fmt) {
    var o = {
        "M+": this.getMonth() + 1, //月份
        "d+": this.getDate(), //日
        "h+": this.getHours(), //小时
        "m+": this.getMinutes(), //分
        "s+": this.getSeconds(), //秒
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度
        "S": this.getMilliseconds() //毫秒
    };
    if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o)
        if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    return fmt;
}
// dict 转 xlsobj

function changeXlsxObj(dataformat){
    var objlst=[]
    var sheetIndex=0
   
    for(var sheet in dataformat){
        var obj= new Array()
        var i=1
        var successTotal=0
        var failedTotal=0
        var db_md_os = {
            dbs:[
                ['MySQL',"Redis",'MSSQL','Oracle'],0,0
            ],
            mds:[
                ['Weblogic','MQ','Apache','Nginx','Tomcat','Tuxedo'],0,0
            ],
            oss:[['Suse','Windows'],0,0]
        }
        obj[0]=[sheet]
        obj[1]=["设备类型","合规数",'不合规数','总数','合规率','不合规率']
        for(var device in dataformat[sheet]){
            i=i+1;
            
            var successNum=dataformat[sheet][device]['success']
            var failedNum = dataformat[sheet][device]['failed']
            var totalNum = successNum+ failedNum
            for(var t in db_md_os){
                // console.log(t[0]+"---")
                if(db_md_os[t][0].includes(device)){
            
                    db_md_os[t][1]=db_md_os[t][1]+successNum
                    db_md_os[t][2]=db_md_os[t][2]+failedNum
                }
            }
            successTotal=successTotal+successNum

            failedTotal=failedTotal+failedNum
            if(totalNum===0){
                obj[i]=[device,successNum,failedNum,totalNum,'-%','-%']
            }else{
                obj[i]=[device,successNum,failedNum,totalNum,(successNum*100/(totalNum*1.0)).toFixed(2)+'%',(failedNum*100/(totalNum*1.0)).toFixed(2)+'%']
            }
           
        }
        var total= successTotal+failedTotal
        // if(total===0){
        //     obj[i+1]=['汇总',successTotal,failedTotal,total,"-%",'-%']
        // }else{
        //     obj[i+1]=['汇总',successTotal,failedTotal,total,(successTotal*100/(total*1.0)).toFixed(2)+"%",(failedTotal*100/(total*1.0)).toFixed(2)+'%']
        // }
       md_total=db_md_os['mds'][1]+db_md_os['mds'][2]
       db_total=db_md_os['dbs'][1]+db_md_os['dbs'][2]
       os_tatal=db_md_os['oss'][1]+db_md_os['oss'][2]
       if(md_total===0){
        obj[i+2]=['中间件合规率',0,0,0,"-%"]
       }else{
        obj[i+2]=['中间件合规率',db_md_os['mds'][1],db_md_os['mds'][2],md_total,(db_md_os['mds'][1]*100/(md_total*1.0)).toFixed(2)+"%"]
       }
       if(db_total===0){
        obj[i+3]=['数据库合规率',0,0,0,"-%"]
       }else{
        
        obj[i+3]=['数据库合规率',db_md_os['dbs'][1],db_md_os['dbs'][2],db_total,(db_md_os['dbs'][1]*100/(db_total*1.0)).toFixed(2)+"%"]
       }
       if(0===os_tatal){
        obj[i+4]=['主机合规率',0,0,0,"-%"]
       }else{
        obj[i+4]=['主机合规率',db_md_os['os'][1],db_md_os['os'][2],os_tatal,(db_md_os['oss'][1]*100/(os_tatal*1.0)).toFixed(2)+"%"]
       }
        
       
        objlst[sheetIndex]={
            name:sheet,
            data:obj
        }
        sheetIndex=sheetIndex+1
    }
    //排序并扩充
    /**
     * 
     * [
     * {
     *  name:'sheetname',
     *  data:[
     *  [],[]
     * ]
     * },
     * ]
     */
    
    var sheetlen= objlst.length
    for(var i=0;i<sheetlen;i++){
        var tmpdata=[]
        tmpdata[0]=objlst[i]['data'][0]
        tmpdata[1]=objlst[i]['data'][1]
        var rows=objlst[i]['data'].length
        // var row=2
        for(var t in deviceType){
            var sign=false
            // console.log(t)
            var dataarr = objlst[i]['data'].slice(2,rows-3)
            for(var item in  dataarr){//获取每一行数据
                console.log(dataarr[item])
                if(dataarr[item].includes(t)){
                    tmpdata.push(dataarr[item])
                    sign=true
                    break
                }
            }
            if(!sign){
                tmpdata.push([t,0,0,0,'0.00%','0.00%'])
            }
        }
        tmpdata.push(objlst[i]['data'][rows-4])
        tmpdata.push(objlst[i]['data'][rows-3])
        tmpdata.push(objlst[i]['data'][rows-2])
        tmpdata.push(objlst[i]['data'][rows-1])
        objlst[i]['data']=tmpdata
    }
  
    return objlst
}
//数据导出
function report(data){
    var ds = new Date()
    var filename = "基线处理结果"+ds.format("yyyyMMdd_hhmmss")+".xlsx"
    console.log(filename)
    const res = xlsx.build(data)
    var dirname=path.dirname(__dirname)
    dirname = path.dirname(dirname)
    // dirname = path.dirname(dirname)
   
    console.log(dirname)
    filename =  path.join(dirname,filename)
    fs.writeFile(filename,res,err=>{
        if(err){
            showinfo.innerHTML=showinfo.innerHTML+"<br>[Error] 导出文件失败！"+err
            return 
        }
        showinfo.innerHTML=showinfo.innerHTML+"<br>[Info] 文件导出成功："+filename
    })

}