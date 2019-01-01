var roc = roc || {};
roc.db = roc.db ||{};
//创建一个连接
roc.db.createDb = function(){
  var conn = new ActiveXObject("ADODB.Connection"), 
    fso = new ActiveXObject("Scripting.FileSystemObject"),
    connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fso.GetFile("./db/message.mdb");
  conn.Open(connstr);//打开数据库 
  roc.db.conn = conn;
  return roc.db.conn;
};
//获取连接
roc.db.getDb = function(){
  if( roc.db.conn ){
    return roc.db.conn;
  }else{
    return roc.db.createDb();
  }
};
//关闭连接
roc.db.closeConn = function(){
  if( roc.db.conn ){
    roc.db.conn.close();
    roc.db.conn = null;
  }
};
//获取结果集
roc.db.getRs = function( sqlStr ){
  var mysql = roc.dom.trim( sqlStr );
  if(mysql == ''){return;}
  var rs = new ActiveXObject("ADODB.Recordset"),
    myConn = roc.db.getDb();
  rs.open( sqlStr , myConn ); 
  return rs; 
};
//关闭结果集
roc.db.closeRs = function( rs ){
  rs.close();
  rs =null;
};
//更新、插入
roc.db.execute = function( sqlStr ){
  var myConn = roc.db.getDb();
  myConn.execute( sqlStr );
  roc.db.closeConn();
};
/*---------Sigma:“我任你践踏我的尊严而毫不生气，是因为我爱你。”---------*/
roc.dom = roc.dom ||{};
roc.dom.id = function( id ){
  if(typeof id == 'string' || id instanceof String) {
    return document.getElementById(id);
  } else if(id && id.nodeName && (id.nodeType == 1 || id.nodeType == 9)) {
    return id;
  }
  return null;  
};
/**
 * @method tagName 根据标签获取指定dom元素
 * @param {String} tagName 元素标签名称
 * @param {HTMLElement} el 元素所属的文档对象 默认为当前文档
 * @return {HTMLElement} 返回HTMLElement元素
 */
roc.dom.tagName = function(tagName, el) {
  var el = el || document;
  return el.getElementsByTagName(tagName);
};
//删除左右两端的空格  
roc.dom.trim = function (str) {   
   return (str+'').replace(/(^\s*)|(\s*$)/g, "");
}
/**
 * @method show 显示目标元素
 * @param {Element} element 目标元素或目标元素的id
 * @param {String} element 目标元素
 */
roc.dom.show = function (element) {
  element = roc.dom.id(element);
  element.style.display = '';
  return element;
};
/**
 * @method hide 隐藏目标元素
 * @param {Element} element 目标元素或目标元素的id
 * @param {String} element 目标元素
 */
roc.dom.hide = function (element) {
  element = roc.dom.id(element);
  element.style.display = 'none';
  return element;
};
/**
 * @method hasClass 判断元素是否含有 class
 * @param {Element} el 元素
 * @param {String} className class 名称
 */
roc.dom.hasClass = function(el, className){
  var re = new RegExp('(^|\\s)' + className + '(\\s|$)');
  return re.test(el.className);
};
/**
 * @method addClass 给元素添加 class 
 * @param {Element} el 元素
 * @param {String} className class 名称
 */
roc.dom.addClass = function(el, className){
  if(!roc.dom.hasClass(el, className)){
    el.className = el.className + ' ' + className;
  }
};
/**
 * @method removeClass 给元素移除 class
 * @param {Element} el 元素
 * @param {String} className class 名称
 */
roc.dom.removeClass = function(el, className){
  var re = new RegExp('(^|\\s)' + className + '(?:\\s|$)')
  el.className = el.className.replace(re, '$1');
};
/**
* date对象命名空间
* 
* @namespace
* @name data
*/
roc.date = roc.date || {};
/**
 * @method format 对目标日期对象进行格式化
 * @param {Object} timestamp 目标日期对象
 * @return {String} str 格式化后的时间
 */
roc.date.format = function(timestamp) {
  if(timestamp =='' )return '';
  var str = '',
    temptime = new Date(Number(timestamp));
  str += temptime.getFullYear() + '-';
  str += temptime.getMonth() + 1 + '-';
  str += temptime.getDate() + ' ';
  str += String(temptime.getHours()).length > 1 ? (temptime.getHours() + ':') : ('0' + temptime.getHours() + ':');
  str += String(temptime.getMinutes()).length > 1 ? (temptime.getMinutes()) : ('0' + temptime.getMinutes());
  return str;
};
/**
 * cookie对象命名空间
 * 
 * @namespace
 * @name cookie
 */
roc.cookie = roc.cookie || {};
/**
 * @method set
 * @param {String} name cookie的键
 * @param {String} value cookie的值
 * @param {String} expires 失效时间(小时)
 * @param {String} domain domain域 
 * @param {String} path 路径 
 * @param {String} secure 是否支持https 
 */
roc.cookie.set = function(name, value, expires, domain, path, secure) {
  var text = encodeURIComponent(value), date = expires;
  if(date && typeof date === 'number') {
    date = new Date();
    date.setTime(date.getTime() + (expires * 3600000));
  }
  if(date instanceof Date) {
    text += '; expires=' + date.toUTCString();
  }
  if(domain) {
    text += '; domain=' + domain;
  }
  if(path) {
    text += '; path=/' + path;
  } else {
    text += '; path=/';
  }
  if(secure) {
    text += '; secure';
  }
  document.cookie = name + '=' + text;
};
/**
 * @method get
 * @param {String} name cookie的键 
 */
roc.cookie.get = function(name) {
  var ret,
    m;
  if(name) {
    if((m = document.cookie.match('(?:^| )' + name + '(?:(?:=([^;]*))|;|$)'))) {
      ret = m[1] ? decodeURIComponent(m[1]) : '';
    }
  }
  return ret;
};
roc.util = roc.util || {};
roc.util.loger = function( type ,msg ){
  switch(type){
    case 'pop': 
      alert(msg);
      break;
    case 'float':break;
    default:break;
  }
};
roc.util.resultBlink = function( msg ){
  //操作闪烁提示
  var $ = roc ,
    opt = $.dom.id("optTip");
  $.util.toogle = $.util.toogle || 0;
  clearTimeout(roc.util.t);//调试
  opt.innerHTML = msg ;
  $.dom.show(opt);
  opt.className = "blink" + $.util.toogle%2;
  $.util.toogle++;
  roc.util.t = setTimeout(function(){
    $.dom.hide(opt);
  },$.config.BLINK_DELAY);
};
roc.util.onlyInputNumber = function( id ){
  //限制文本框、文本域只能输入数字
  var $ = roc ,
    num = $.dom.id( id );
  if( num.tagName.toLowerCase() != 'input' || num.tagName.toLowerCase() != 'textarea' ){
    return ;
  }
  $.util.addEvent( num , 'keypress' , function( e ){
    var e = e || window.event ;
    if(e.keyCode >= 48 && e.keyCode <= 57){alert()
      return true;
    }
    return false;
  });
};
roc.util.addEvent = function(elem, type, fn, useCapture) {
  if(elem.addEventListener) { //DOM2.0
    elem.addEventListener(type, fn, useCapture);
    return true;
  } else if(elem.attachEvent) { //IE5+
    elem.attachEvent('on' + type, fn);
    return true;
  } else { //DOM 0
    elem['on' + type] = fn;
  }
};
roc.config = roc.config || {};
roc.config = roc.config ||{
  BLINK_DELAY:3000,
  SELECT_DELAY:1000
}
roc.search = roc.search || {};
 roc.search.getValues = function( e ){
    //批量获取表单值,用于插入
    var $ = roc,
      allIsNull = true,
    wrapStr = function( num ){
      return '"'+ num + '"';
    },
    vals =[];
  for(var i in e[0]){
    var v =$.dom.trim($.dom.id( e[0][i] ).value + '');
    if( v != ''){
      allIsNull = false;
    }
    switch( e[1][i] ){
      case 'date':
      case 'text':
        vals.push( wrapStr(v) );
        break;
      case 'num':
        vals.push( v );
        break;
      default:break;
    }
  }
  if(allIsNull){
    return false;
  }
  return vals.join(',');
};
roc.search.getSelSql = function(){
  //组装搜索sql
  var $ = roc,
    addr = $.dom.trim($.dom.id("s_uaddr").value),
    phone = $.dom.trim($.dom.id("s_uphone").value),
    style = $.dom.trim($.dom.id("s_style").value),
    year = $.dom.trim($.dom.id("s_year").value ),
    month = $.dom.trim($.dom.id("s_month").value),
    date = $.dom.trim($.dom.id("s_date").value),
    datetype = $.dom.trim($.dom.id("s_datetype").value),
    mysql = 'select * from inslist where 1=1 ',
    datetypeName = datetype == 0 ? 'selltime':'addtime';
  if( addr != ''){
    mysql += ' and uaddr like "%' + addr + '%"';  
  }
  if( phone != ''){
    mysql += ' and uphone ="' + phone + '"';
  }
  if( style != ''){
    mysql += ' and typeid = ' + style + '';
  }
  if( year !=''){
    mysql += ' and year(' + datetypeName + ') = ' + year + '';
  }
  if( month !=''){
    mysql += ' and month(' + datetypeName + ') = ' + month + '';
  }
  if( date !=''){
    mysql += ' and date(' + datetypeName + ') = ' + date + '';
  }
  return mysql;
};
//搜索
roc.search.seeking = function(){
  if( !roc.search.getLock()){return;}
  var $ = roc ,
    mySql = $.search.getSelSql();
    html = $.search.getSel( mySql);
  $.search.setLock(false);
  $.dom.id("searchResult").innerHTML = html;
  $.util.resultBlink("查询完毕");//闪烁
};
roc.search.getSel = function( sqlStr ){
  //查询
  var $ = roc,
    rs = $.db.getRs( sqlStr ),
    filtRs = function ( str ){//处理字段
      return ( str + '' ) == 'null' ? '':str;
  },
  num = 1;
  total_receive = 0,
  total_prize = 0,
  html = "<table class='list' id='memoryDetails'>"
      + "<colgroup>"
      + "<col class='pid' />"
      + "<col class='uaddr' />"
      + "<col class='phone' />"
      + "<col class='number'/>"
      + "<col class='money' span='3'/>"
      + "<col class='number' />"
      + "<col class='phone' />"
      + "<col class='number' />"
      + "<col class='date' />"
      + "<col class='date' />"
      + "</colgroup>"
      + "<tr class='secondRow doNotFilter'>"
      + " <th class='pid'> 序号 </th>"
      + " <th class='uaddr'> 用户地址 </th>"
      + " <th class='phone'> 用户电话 </th>"
      + " <th class='number'> 型号 </th>"
      + " <th class='money'> 代收款 </th>"
      + " <th class='money'> 货款 </th>"
      + " <th class='money'> 余额 </th>"
      + " <th class='number'> 安装人 </th>"
      + " <th class='phone'> 销售电话 </th>"
      + " <th class='number'> 备注 </th>"
      + " <th class='date'> 销售日期 </th>"
      + " <th class='date'> 记录时间 </th>"
      + "</tr>";
  while(!rs.EOF) 
  {
    var id   = num ,//filtRs(rs.Fields("id") ),
    uaddr    = filtRs(rs.Fields("uaddr") ),
    uphone   = filtRs(rs.Fields("uphone") ),
    typeid   = filtRs(rs.Fields("typeid") ),
    received  = filtRs(rs.Fields("received") ),
    prize    = filtRs(rs.Fields("prize") ),
    unreceived = filtRs(rs.Fields("unreceived") ),
    installerid = filtRs(rs.Fields("installerid") ),
    sellerid  = filtRs(rs.Fields("sellerid") ),
    remark   = filtRs(rs.Fields("remark") ),
    selltime  = $.date.format(filtRs(rs.Fields("selltime") )),
    addtime   = $.date.format(filtRs(rs.Fields("addtime") ) );
    html += "<tr jsselect='browzr_data'>"
    +"<td class='pid'>" + id + "</td>"
    +"<td class='uaddr'>" + uaddr + "</td>"
    +"<td class='phone'>" + uphone + "</td>"
    +"<td class='number'>" + typeid + "</td>"
    +"<td class='money'>" + received + "</td>"
    +"<td class='money'>" + prize + "</td>"
    +"<td class='money'>" + unreceived + "</td>"
    +"<td class='number'>" + installerid + "</td>"
    +"<td class='phone'>" + sellerid + "</td>"
    +"<td class='number'>" + remark + "</td>"
    +"<td class='date'>" + selltime + "</td>"
    +"<td class='date'>" + addtime + "</td>"
    +"</tr>";
    //统计项
    total_receive += received,
    total_prize += prize,
    num++;
    rs.moveNext();
  } 
  html = html 
  +"<tr class='total doNotFilter'>"
  +"<td class='pid'></td>"
  +"<td class='uaddr'>Σ </td>"
  +"<td class='number'></td>"
  +"<td class='number'></td>"
  +"<td class='number'>" + total_receive +"</td>"
  +"<td class='number'>" + total_prize +"</td>"
  +"<td class='number'>" + (total_prize - total_receive ) +"</td>"
  +"<td class='number'></td>"
  +"<td class='number'></td>"
  +"<td class='number'></td>"
  +"<td class='date'></td>"
  +"<td class='date'></td>"
  +"</table>";
  $.db.closeRs(rs);
  $.db.closeConn(); 
  return html;
};
roc.search.getLock = function(){
  //查询锁
  if( typeof roc.search.searchLock == 'undefined' ){
    roc.search.setLock(false);
  }
  return roc.search.searchLock;
};
roc.search.setLock = function( key ){
  roc.search.searchLock = key;
};
//[[id],[type]]
roc.search.addEls = [[
        "uaddr",
        "uphone",
        "typeid",
        "received",
        "prize",
        "unreceived",
        "installerid",
        "sellerid",
        "remark",
        "selltime"
        ],[
        'text',
        'text',
        'num',
        'num',
        'num',
        'num',
        'num',
        'num',
        'text',
        'date'
        ]];
roc.search.insert = function(){
  //插入安装单记录
  var $ = roc,
    getV = $.search.getValues( $.search.addEls );
  if(!getV){
    $.util.loger('pop','请填写信息后再保存！');
    return; 
  }
  var sqlStr = 'insert into inslist (uaddr,uphone,typeid,received,prize,unreceived,installerid,sellerid,remark,selltime) values ('+ getV +')';
  $.db.execute( sqlStr );
  $.util.resultBlink('保存安装单成功');
};
/*显示与隐藏*/
roc.dom.switchDiv = function( objDiv){
  var $ = roc ,
    cookieName = objDiv.id + 'cookie';
  if( objDiv.style.display =='' || objDiv.style.display =='none' ){
    $.dom.show( objDiv );
    $.cookie.set(cookieName,0,9999999);
  }else{ 
    $.dom.hide( objDiv );
    $.cookie.set(cookieName,1,9999999);
  }
};
//货物型号操作
roc.tstyle = roc.tstyle || {};
roc.tstyle.els = [
          ['tname','tprize','tdesc'],
          ['text','text','text']
        ];
roc.tstyle.insert = function(){
  //插入记录
  var $ = roc,
    getV = $.search.getValues( $.tstyle.els );
  if( !getV ){
    $.util.loger('pop','请填写信息后再保存！');
    return; 
  }
  var sqlStr = 'insert into type ( tname , tprize , tdesc ) values ('+ getV +')';
  //$.util.loger('pop',sqlStr);
  $.db.execute( sqlStr );
  $.util.resultBlink('保存成功！');
  $.util.flushInput($.tstyle.els);
};
roc.util.flushInput = function( els ){
  var $ = roc ;
  for(var i = 0 ; i < els.length ; i ++){
    var e = $.dom.id(els[i]+'');
    /* if(e.tagName == 'input' && e.type =='text'){*/
      e.value = '';
    /*}*/
  }
};
//type{id,tname,tprize}
roc.tstyle.getStyle = function( optId ){
  //获取类型列表
  var $ = roc ,
    mySql = 'select * from type where isdel = 0 ',
    rs = $.db.getRs( mySql ),
    filtRs = function ( str ){//处理字段
      return (str+'')=='null' ? '':str;
  },
  myOpt = $.dom.id( optId ),
  optIndex = 1;
  while(! rs.EOF ){
    var id = filtRs(rs.Fields('id')),
      prize = filtRs(rs.Fields('tprize')),
      name = filtRs(rs.Fields('tname'));
      desc = filtRs(rs.Fields('tdesc'));
    myOpt.options[optIndex] = new Option( name , id );
    myOpt.options[optIndex].title = '价格：' + prize + ' | 描述:' + desc;
    optIndex++;
    rs.moveNext();
  }
  $.db.closeRs(rs);
  $.db.closeConn(); 
};
;(function(){
  var $ = roc ;
  $.dom.id("save").onclick = function(){
    //保存
    $.search.insert();
  }
  $.dom.id("searchBtn").onclick = function(){
    //提检
    $.search.seeking();
  }
  //初始化查询安装单 年
  for(var i = 0 ; i <= 10 ; i++ ){
    $.dom.id("s_year").options[i] = new Option(2010 + i , 2010 + i );
    if( 2010+i+'' == (new Date()).getYear() ){
      $.dom.id("s_year").options[i].selected = true;
    }
  }
  //初始化查询安装单月份
  for(var i = 1 ; i <= 12 ; i++ ){
    $.dom.id("s_month").options[i] = new Option(i,i);
  }
  //提检条件字段id 修改触发查询
  $.dom.s_fields = ["s_uaddr","s_uphone","s_style","s_datetype",'s_year','s_month','s_date'];
  for(var i = 0 ; i < $.dom.s_fields.length ; i ++){
    var f = $.dom.s_fields[i];
    $.dom.id(f).onpropertychange = function(){
      if(event.propertyName == 'value'){
        $.search.setLock(true);
        if($.search.t){
          clearTimeout($.search.t);
        }
        $.search.t = setTimeout(function(){
          $.search.seeking();
        },$.config.SELECT_DELAY);
      }
    }
    $.dom.id(f).onfocus = function(){
      $.dom.addClass(this,"focusit");
    };
    $.dom.id(f).onblur = function(){
      $.dom.removeClass(this,"focusit");
    };
  }
  $.dom.id('saveType').onclick = function(){
    //货物类型
    $.tstyle.insert();
  };
  //取出类型列表
  $.tstyle.getStyle('typeid');
  $.tstyle.getStyle('s_style');
/* //$.dom.id("s_uaddr").onkeyup = $.dom.id("s_uphone").onkeyup = $.dom.id("s_style").onkeyup = function(){
  $.dom.id("s_uaddr").onblur = $.dom.id("s_uphone").onblur = $.dom.id("s_style").onblur = function(){
    $.dom.removeClass(this,"focusit");
  }
  $.dom.id("s_uaddr").onfocus = $.dom.id("s_uphone").onfocus = $.dom.id("s_style").onfocus = function(){
    $.dom.addClass(this,"focusit");
  }*/
  //导航样式切换
  for(var i = 0 ; i < $.search.addEls.length ; i++ ){
    var curObj = $.dom.id($.search.addEls[0][i]+'');
    curObj.onfocus = function(){
      $.dom.addClass(this,'focusit');
    }
    curObj.onblur = function(){
      $.dom.removeClass(this,'focusit');
    }
  }
  //添加导航点击事件
  var lis = $.dom.tagName('li',$.dom.id("ulNav"));
  for(var i = 0 ; i < lis.length ; i ++ ){
    $.dom.hide( $.dom.id(lis[i].id + 'Div'));
    lis[i].onclick = function(){
    for(var n = 0 ; n < lis.length ; n ++ ){
      $.dom.removeClass(lis[n],'click');
      $.dom.hide( $.dom.id(lis[n].id + 'Div'));
    }
    $.dom.show( $.dom.id(this.id + "Div"));
    $.dom.addClass(this,"click");
    $.cookie.set('showWhichDiv', this.id);
  }
  }
  //默认的载入显示页面
  var showWhichDiv = $.cookie.get("showWhichDiv") || "searchList";
  $.dom.addClass($.dom.id(showWhichDiv),"click");
  $.dom.show($.dom.id(showWhichDiv + 'Div'));
  //日期控件,感谢此控件开发者的分享，祝你有个好女朋友！
  J('#selltime').calendar({ format:'yyyy-MM-dd HH:mm:ss' });
  var numFields = ['s_uphone','s_date','uphone','received','prize','unreceived','installerid'];
  for( var i = 0 ; i < numFields.length ; i ++ ){
    $.util.onlyInputNumber( numFields[i] );
  }
})();
