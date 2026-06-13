// ============================================================
// KIMMIZU - Backend API v3.0 — Tối ưu tốc độ + bảo mật
// ============================================================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  NHAN_VIEN:'NhanVien', CA_LAM_VIEC:'CaLamViec', VIEC_CA:'DanhSachViecCa',
  CONG_VIEC:'CongViec', PHAN_CONG:'PhanCong', THONG_BAO:'ThongBao',
  KPI:'KPI', CAI_DAT:'CaiDat', NHAT_KY:'NhatKy', DA_DOC:'DaDocThongBao',
};
const TRANG_THAI = {CHUA_LAM:'Chưa làm',DANG_LAM:'Đang làm',HOAN_THANH:'Hoàn thành',TRA_VIEC:'Trả việc'};
const LOAI_VIEC  = {CA:'Ca',DINH_KY:'Định kỳ',PHAT_SINH:'Phát sinh'};

// ── Chỉ log các write actions — KHÔNG log read (tăng tốc ~150ms/request) ──
const WRITE_ACTIONS = new Set([
  'login','dangXuat','nhanViec','hoanThanhViec','traViec','taoViecPhatSinh',
  'taoCongViec','suaCongViec','xoaCongViec','taoThongBao','suaThongBao',
  'xoaThongBao','docThongBao','luuNhanVien','luuCaiDat','nhacNhoThongBao',
]);

// ── ENTRY POINT ──
function doGet(e) {
  const p=e.parameter, action=p.action||'', token=p.token||'';
  if(action==='login') return handleLogin(p);
  if(action==='ping')  return json({success:true,time:fmtDateTime(new Date())});
  const user=verifyToken(token);
  if(!user) return json({success:false,error:'Unauthorized'});
  try{
    const result=dispatch(action,p,user);
    // Chỉ log write actions (không log read để tăng tốc)
    if(WRITE_ACTIONS.has(action)){
      const sp={};for(const k in p){if(k!=='token')sp[k]=p[k];}
      ghiNhatKy(getLoaiLog(action),action,user.tenNV,JSON.stringify(sp).substring(0,120));
    }
    return result;
  }catch(err){
    ghiNhatKy('LỖI',action,user.tenNV,'Lỗi: '+err.toString().substring(0,200));
    return json({success:false,error:'Lỗi hệ thống. Thử lại sau.'});
  }
}
function doPost(e){return doGet(e);} // Xử lý chung
function doOptions(e){return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);}

function dispatch(action,p,user){
  switch(action){
    case 'getCongViec':     return getCongViec(p,user);
    case 'nhanViec':        return nhanViec(p,user);
    case 'hoanThanhViec':   return hoanThanhViec(p,user);
    case 'traViec':         return traViec(p,user);
    case 'taoViecPhatSinh': return taoViecPhatSinh(p,user);
    case 'taoCongViec':     return taoCongViec(p,user);
    case 'suaCongViec':     return suaCongViec(p,user);
    case 'xoaCongViec':     return xoaCongViec(p,user);
    case 'getThongBao':     return getThongBao(p,user);
    case 'getThongBaoAdmin':return getThongBaoAdmin(p,user);
    case 'nhacNhoThongBao': return nhacNhoThongBao(p,user);
    case 'docThongBao':     return docThongBao(p,user);
    case 'taoThongBao':     return taoThongBao(p,user);
    case 'suaThongBao':     return suaThongBao(p,user);
    case 'xoaThongBao':     return xoaThongBao_TB(p,user);
    case 'getKPI':          return getKPI(p,user);
    case 'getDashboard':    return getDashboard(p,user);
    case 'getNhanVien':     return getNhanVien(p,user);
    case 'luuNhanVien':     return luuNhanVien(p,user);
    case 'getBaoCao':       return getBaoCao(p,user);
    case 'getCaiDat':       return getCaiDat(p,user);
    case 'luuCaiDat':       return luuCaiDat(p,user);
    case 'dangXuat':        return dangXuat(p,user);
    case 'getNhatKy':       return getNhatKy(p,user);
    case 'xuatBaoCaoExcel':  return xuatBaoCaoExcel(p,user);
    default: return json({success:false,error:'Action không hợp lệ'});
  }
}

// ── XÁC THỰC + BẢO MẬT ──
function handleLogin(p) {
  const tenNV=(p.tenNV||'').trim().toLowerCase(), mk=(p.matKhau||'').trim();
  if(!tenNV||!mk) return json({success:false,error:'Thiếu thông tin đăng nhập'});
  if(mk.length>20) return json({success:false,error:'Thông tin không hợp lệ'});

  const sh=getSheet(SHEETS.NHAN_VIEN), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    const row=data[i]; if(!row[0]) continue;
    if(String(row[1]||'').trim().toLowerCase()===tenNV &&
       String(row[2]||'').trim()===mk &&
       String(row[5]||'').trim()==='Hoạt động'){
      const token=Utilities.getUuid(), expiry=new Date(Date.now()+30*24*3600*1000);
      // Batch write token + expiry
      sh.getRange(i+1,7,1,2).setValues([[token,expiry.toISOString()]]);
      ghiNhatKy('ĐĂNG NHẬP','login',String(row[1]).trim(),'OK');
      return json({success:true,user:{
        id:row[0], tenNV:String(row[1]).trim(), vaiTro:String(row[3]).trim(),
        email:String(row[4]||''), token, expiry:expiry.toISOString()
      }});
    }
  }
  ghiNhatKy('LOGIN THẤT BẠI','login',tenNV,'Sai thông tin');
  return json({success:false,error:'Tên đăng nhập hoặc mật khẩu không đúng'});
}

function verifyToken(token) {
  if(!token||token.length<10) return null;
  const data=getSheet(SHEETS.NHAN_VIEN).getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    const row=data[i];
    if(String(row[6]||'')===token){
      if(new Date(row[7])>new Date())
        return {id:row[0],tenNV:String(row[1]).trim(),vaiTro:String(row[3]).trim(),email:String(row[4]||'')};
      return null; // hết hạn — không log để tránh spam
    }
  }
  return null;
}

// ── NHẬT KÝ (chỉ write actions) ──
let _nkSheet=null;
function ghiNhatKy(loai,action,nguoi,ghiChu) {
  try{
    if(!_nkSheet){
      _nkSheet=SS.getSheetByName(SHEETS.NHAT_KY);
      if(!_nkSheet){
        _nkSheet=SS.insertSheet(SHEETS.NHAT_KY);
        _nkSheet.getRange(1,1,1,5).setValues([['ThoiGian','Loai','Action','NguoiDung','GhiChu']]);
        styleHeader(_nkSheet,5); _nkSheet.setFrozenRows(1);
      }
    }
    _nkSheet.appendRow([new Date(),loai||'',action||'',nguoi||'',ghiChu||'']);
  }catch(e){_nkSheet=null;}
}
function getNhatKy(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const data=getSheet(SHEETS.NHAT_KY).getDataRange().getValues();
  const rows=[];
  for(let i=Math.max(1,data.length-300);i<data.length;i++){
    const r=data[i]; if(!r[0]) continue;
    rows.push({thoiGian:r[0] instanceof Date?fmtDateTime(r[0]):String(r[0]),loai:r[1],action:r[2],nguoi:r[3],ghiChu:r[4]});
  }
  rows.reverse();
  return json({success:true,data:rows});
}

// ── CA LÀM VIỆC ──
function getCaHienTai(){
  const ca=(getCaiDatObj().caLamViec)||[
    {ten:'Sáng',batDau:'07:00',ketThuc:'10:00'},{ten:'Trưa',batDau:'10:00',ketThuc:'14:00'},
    {ten:'Chiều',batDau:'14:00',ketThuc:'18:00'},{ten:'Tối',batDau:'18:00',ketThuc:'22:00'}
  ];
  const now=new Date(),phut=now.getHours()*60+now.getMinutes();
  let caGan=ca[0],minDiff=Infinity;
  for(const c of ca){
    const [bH,bM]=c.batDau.split(':').map(Number),[kH,kM]=c.ketThuc.split(':').map(Number);
    if(phut>=bH*60+bM && phut<kH*60+kM) return c.ten;
    const diff=Math.abs(phut-(bH*60+bM));if(diff<minDiff){minDiff=diff;caGan=c;}
  }
  return caGan.ten;
}

// ── CÔNG VIỆC - ĐỌC (tối ưu: 1 pass, không loop thừa) ──
function getCongViec(p,user){
  const ngay=p.ngay||fmtDate(new Date()), loai=p.loai||'', ca=p.ca||'', tatCa=(ca==='all');
  const caHT=getCaHienTai();
  const data=getSheet(SHEETS.CONG_VIEC).getDataRange().getValues();
  const result=[], s={chuaLam:0,dangLam:0,hoanThanh:0,traViec:0};
  for(let i=1;i<data.length;i++){
    const row=data[i]; if(!row[0]) continue;
    if(readDateFromSheet(row[5])!==ngay) continue;
    const rLoai=String(row[3]||'').trim(), rCa=String(row[4]||'').trim(), rTT=String(row[6]||'').trim();
    if(loai && rLoai!==loai) continue;
    if(!tatCa){
      if(user.vaiTro==='Nhân viên' && rLoai===LOAI_VIEC.CA){
        if(rCa!==(ca||caHT)) continue;
      } else if(ca && rLoai===LOAI_VIEC.CA && rCa!==ca) continue;
    }
    if(rTT===TRANG_THAI.CHUA_LAM)        s.chuaLam++;
    else if(rTT===TRANG_THAI.DANG_LAM)   s.dangLam++;
    else if(rTT===TRANG_THAI.HOAN_THANH) s.hoanThanh++;
    else if(rTT===TRANG_THAI.TRA_VIEC)   s.traViec++;
    result.push(rowToViec(row));
  }
  return json({success:true,data:result,stats:s,caHienTai:caHT});
}
function rowToViec(row){
  return {id:String(row[0]),tenViec:row[1],moTa:row[2],loaiViec:row[3],ca:row[4],
    ngayThucHien:readDateFromSheet(row[5]),trangThai:row[6],nguoiNhan:row[7],
    thoiGianNhan:row[8]?fmtDateTime(new Date(row[8])):'',
    thoiGianHT:row[9]?fmtDateTime(new Date(row[9])):'',
    uuTien:row[10],hanChot:readDateFromSheet(row[11]),chuKy:row[12],
    ngayTao:readDateFromSheet(row[13]),nguoiTao:row[14]};
}

// ── CÔNG VIỆC - GHI ──
function taoViecPhatSinh(p,user){
  const ten=(p.tenViec||'').trim();
  if(!ten) return json({success:false,error:'Thiếu tên công việc'});
  const id='CV_'+Date.now(), hom=fmtDate(new Date());
  writeViec([id,ten,p.moTa||'',LOAI_VIEC.PHAT_SINH,'',hom,TRANG_THAI.CHUA_LAM,'','','',p.uuTien||'Bình thường',p.hanChot||'','',hom,user.tenNV]);
  return json({success:true,id});
}
function taoCongViec(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const ten=(p.tenViec||'').trim(); if(!ten) return json({success:false,error:'Thiếu tên'});
  const loai=p.loaiViec||LOAI_VIEC.CA, ca=(p.ca||'').trim();
  const id='CV_'+Date.now(), hom=fmtDate(new Date()), ngay=p.ngayThucHien||hom;
  writeViec([id,ten,p.moTa||'',loai,ca,ngay,TRANG_THAI.CHUA_LAM,'','','',p.uuTien||'Bình thường',p.hanChot||'',p.chuKy||'',hom,user.tenNV]);
  return json({success:true,id,debug:{loai,ca,ngay}});
}
function suaCongViec(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.CONG_VIEC), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  const map={tenViec:2,moTa:3,loaiViec:4,ca:5,ngayThucHien:6,uuTien:11,hanChot:12,chuKy:13};
  for(const [k,col] of Object.entries(map)){if(p[k]!==undefined && p[k]!=='')sh.getRange(ri,col).setValue(p[k]);}
  return json({success:true});
}
function xoaCongViec(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.CONG_VIEC), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  sh.deleteRow(ri);
  return json({success:true});
}

// ── NHẬN / HOÀN THÀNH / TRẢ (batch writes) ──
function nhanViec(p,user){
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.CONG_VIEC), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  const row=sh.getRange(ri,1,1,15).getValues()[0];
  if(row[6]!==TRANG_THAI.CHUA_LAM) return json({success:false,error:'Việc không còn ở trạng thái Chưa làm'});
  const now=new Date();
  sh.getRange(ri,7,1,3).setValues([[TRANG_THAI.DANG_LAM,user.tenNV,now.toISOString()]]);
  ghiPhanCong(p.id,row[1],user.tenNV,'Nhận việc',now);
  return json({success:true});
}
function hoanThanhViec(p,user){
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.CONG_VIEC), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  const row=sh.getRange(ri,1,1,15).getValues()[0];
  if(row[6]!==TRANG_THAI.DANG_LAM) return json({success:false,error:'Việc không ở trạng thái Đang làm'});
  if(row[7]!==user.tenNV && user.vaiTro!=='Quản lý') return json({success:false,error:'Bạn không phải người nhận việc này'});
  const now=new Date();
  sh.getRange(ri,7).setValue(TRANG_THAI.HOAN_THANH);
  sh.getRange(ri,10).setValue(now.toISOString());
  ghiPhanCong(p.id,row[1],user.tenNV,'Hoàn thành',now);
  capNhatKPI(user.tenNV,row[8]?new Date(row[8]):now,now);
  return json({success:true});
}
function traViec(p,user){
  const lyDo=p.lyDo||'';
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.CONG_VIEC), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  const row=sh.getRange(ri,1,1,15).getValues()[0];
  if(row[6]!==TRANG_THAI.DANG_LAM) return json({success:false,error:'Việc không ở trạng thái Đang làm'});
  if(row[7]!==user.tenNV && user.vaiTro!=='Quản lý') return json({success:false,error:'Bạn không phải người nhận việc này'});
  const now=new Date();
  sh.getRange(ri,7,1,3).setValues([[TRANG_THAI.CHUA_LAM,'','']]);
  ghiPhanCong(p.id,row[1],user.tenNV,'Trả việc: '+lyDo,now);
  return json({success:true});
}

// ── THÔNG BÁO ──
function getThongBao(p,user){
  const data=getSheet(SHEETS.THONG_BAO).getDataRange().getValues();
  const ddSet=getDaDocSet(user.tenNV), result=[];
  for(let i=1;i<data.length;i++){
    const row=data[i]; if(!row[0]) continue;
    if(row[3]!=='Tất cả' && row[3]!==user.tenNV) continue;
    const id=String(row[0]);
    result.push({id,tieuDe:row[1],noiDung:row[2],nguoiNhan:row[3],
      ngayTao:row[4] instanceof Date?fmtDateTime(row[4]):String(row[4]),
      daDoc:ddSet.has(id),nguoiTao:row[6]});
  }
  result.sort((a,b)=>new Date(b.ngayTao)-new Date(a.ngayTao));
  return json({success:true,data:result,chuaDoc:result.filter(t=>!t.daDoc).length});
}
function getDaDocSet(tenNV){
  try{
    const data=getSheet(SHEETS.DA_DOC).getDataRange().getValues();
    const set=new Set();
    for(let i=1;i<data.length;i++){if(data[i][0]&&String(data[i][1])===tenNV)set.add(String(data[i][0]));}
    return set;
  }catch(e){return new Set();}
}
function getDaDocSet_byTBId(tbId){
  try{
    const data=getSheet(SHEETS.DA_DOC).getDataRange().getValues();
    const set=new Set();
    for(let i=1;i<data.length;i++){if(String(data[i][0])===String(tbId))set.add(String(data[i][1]));}
    return set;
  }catch(e){return new Set();}
}
function docThongBao(p,user){
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  if(!getDaDocSet(user.tenNV).has(p.id)) getSheet(SHEETS.DA_DOC).appendRow([p.id,user.tenNV,new Date()]);
  return json({success:true});
}
function taoThongBao(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const td=(p.tieuDe||'').trim(), nd=(p.noiDung||'').trim();
  if(!td||!nd) return json({success:false,error:'Thiếu tiêu đề hoặc nội dung'});
  const id='TB_'+Date.now();
  getSheet(SHEETS.THONG_BAO).appendRow([id,td,nd,p.nguoiNhan||'Tất cả',new Date(),false,user.tenNV]);
  return json({success:true,id});
}
function suaThongBao(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.THONG_BAO), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  if(p.tieuDe) sh.getRange(ri,2).setValue(p.tieuDe);
  if(p.noiDung) sh.getRange(ri,3).setValue(p.noiDung);
  if(p.nguoiNhan) sh.getRange(ri,4).setValue(p.nguoiNhan);
  return json({success:true});
}
function xoaThongBao_TB(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.THONG_BAO), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  sh.deleteRow(ri);
  try{
    const ddSh=getSheet(SHEETS.DA_DOC), ddData=ddSh.getDataRange().getValues();
    for(let i=ddData.length-1;i>=1;i--){if(String(ddData[i][0])===String(p.id))ddSh.deleteRow(i+1);}
  }catch(e){}
  return json({success:true});
}
function getThongBaoAdmin(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const tbData=getSheet(SHEETS.THONG_BAO).getDataRange().getValues();
  const ddData=getSheet(SHEETS.DA_DOC).getDataRange().getValues();
  const nvData=getSheet(SHEETS.NHAN_VIEN).getDataRange().getValues();
  const dsNV=[];
  for(let i=1;i<nvData.length;i++){
    const r=nvData[i]; if(!r[0]) continue;
    if(String(r[3]).trim()==='Nhân viên' && String(r[5]).trim()==='Hoạt động')
      dsNV.push({tenNV:String(r[1]).trim(),email:String(r[4]||'').trim()});
  }
  const ddMap={};
  for(let i=1;i<ddData.length;i++){
    const r=ddData[i]; if(!r[0]) continue;
    const id=String(r[0]); if(!ddMap[id])ddMap[id]=new Set(); ddMap[id].add(String(r[1]));
  }
  const result=[];
  for(let i=1;i<tbData.length;i++){
    const row=tbData[i]; if(!row[0]) continue;
    const id=String(row[0]), nguoiNhan=row[3], ddSet=ddMap[id]||new Set();
    const danhSach=nguoiNhan==='Tất cả'?dsNV.map(n=>n.tenNV):[nguoiNhan];
    const daDoc=danhSach.filter(nv=>ddSet.has(nv));
    const chuaDoc=danhSach.filter(nv=>!ddSet.has(nv));
    result.push({id,tieuDe:row[1],noiDung:row[2],nguoiNhan,
      ngayTao:row[4] instanceof Date?fmtDateTime(row[4]):String(row[4]),nguoiTao:row[6],
      tongNhan:danhSach.length,soDaDoc:daDoc.length,soChuaDoc:chuaDoc.length,
      dsDaDoc:daDoc,dsChuaDoc:chuaDoc,
      emailChuaDoc:chuaDoc.map(nv=>{const f=dsNV.find(n=>n.tenNV===nv);return f?f.email:'';}).filter(e=>e)});
  }
  result.sort((a,b)=>new Date(b.ngayTao)-new Date(a.ngayTao));
  return json({success:true,data:result,dsNV});
}
function nhacNhoThongBao(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  if(!p.id) return json({success:false,error:'Thiếu ID'});
  const sh=getSheet(SHEETS.THONG_BAO), ri=findById(sh,p.id);
  if(!ri) return json({success:false,error:'Không tìm thấy'});
  const row=sh.getRange(ri,1,1,7).getValues()[0];
  const tieuDe=String(row[1]||''), nd=String(row[2]||''), nguoiNhan=String(row[3]||'Tất cả');
  const nvData=getSheet(SHEETS.NHAN_VIEN).getDataRange().getValues();
  const emailMap={}, dsNV=[];
  for(let i=1;i<nvData.length;i++){
    const r=nvData[i]; if(!r[0]) continue;
    emailMap[String(r[1]).trim()]=String(r[4]||'').trim();
    if(String(r[3]).trim()==='Nhân viên' && String(r[5]).trim()==='Hoạt động') dsNV.push(String(r[1]).trim());
  }
  const danhSach=nguoiNhan==='Tất cả'?dsNV:[nguoiNhan];
  const ddSet=getDaDocSet_byTBId(p.id);
  const chuaDoc=danhSach.filter(nv=>!ddSet.has(nv));
  if(!chuaDoc.length) return json({success:false,error:'Tất cả đã đọc rồi!'});
  const sent=[],failed=[],noEmail=[];
  for(const ten of chuaDoc){
    const em=emailMap[ten]||'';
    if(em&&em.includes('@')){
      try{MailApp.sendEmail({to:em,subject:'[KIMMIZU] Nhắc đọc: '+tieuDe,htmlBody:buildEmailHtml(ten,tieuDe,nd)});sent.push(ten);}
      catch(err){failed.push(ten);}
    }else noEmail.push(ten);
  }
  return json({success:true,message:'Đã gửi '+sent.length+' email',sent,failed,noEmail,dsChuaDoc:chuaDoc});
}
function buildEmailHtml(tenNV,tieuDe,nd){
  return '<div style="font-family:Arial,sans-serif;max-width:500px"><div style="background:#3B2314;padding:16px 24px;border-radius:8px 8px 0 0"><h2 style="color:#F5E6C8;margin:0">KIMMIZU</h2></div><div style="background:#FDF8EF;padding:20px;border:1px solid #E8D5A3"><p>Xin chào <strong>'+tenNV+'</strong>,</p><div style="background:#fff;border-left:4px solid #3B2314;padding:14px;margin:12px 0"><p style="font-weight:bold;color:#3B2314;margin:0 0 8px">'+tieuDe+'</p><p style="color:#6B3F26;margin:0">'+nd+'</p></div><p style="color:#A0622A;font-size:13px">Mở app KIMMIZU để đọc và xác nhận.</p></div></div>';
}

// ── KPI ──
function getKPI(p,user){
  const loai=p.loai||'thang', tenNV=(user.vaiTro==='Quản lý'&&p.tenNV)?p.tenNV:user.tenNV;
  const now=new Date(); let startDate,endDate;
  if(p.tuNgay&&p.denNgay){startDate=parseDate(p.tuNgay);endDate=parseDate(p.denNgay);endDate.setHours(23,59,59);}
  else if(loai==='tuan'){const d=now.getDay()||7;startDate=new Date(now);startDate.setDate(now.getDate()-d+1);startDate.setHours(0,0,0,0);endDate=new Date(startDate);endDate.setDate(startDate.getDate()+6);endDate.setHours(23,59,59);}
  else{startDate=new Date(now.getFullYear(),now.getMonth(),1);endDate=new Date(now.getFullYear(),now.getMonth()+1,0,23,59,59);}
  const pcData=getSheet(SHEETS.PHAN_CONG).getDataRange().getValues();
  const ht=[],tv=[];
  for(let i=1;i<pcData.length;i++){
    const r=pcData[i]; if(!r[0]||r[3]!==tenNV) continue;
    const t=new Date(r[5]); if(t<startDate||t>endDate) continue;
    if(r[4]==='Hoàn thành')ht.push(r); if(String(r[4]).startsWith('Trả việc'))tv.push(r);
  }
  let tongPhut=0,demTD=0;
  for(const h of ht){const n=pcData.find(r=>r[1]===h[1]&&r[3]===tenNV&&r[4]==='Nhận việc');if(n){const phut=(new Date(h[5])-new Date(n[5]))/60000;if(phut>0&&phut<480){tongPhut+=phut;demTD++;}}}
  const tocDo=demTD>0?Math.round(tongPhut/demTD):0;
  const ts=(getCaiDatObj().kpiTrongSo)||{hoanThanh:60,traViec:25,tocDo:15};
  const dHT=Math.min(ht.length*5,100)*ts.hoanThanh/100;
  const dTV=Math.max(0,100-tv.length*10)*ts.traViec/100;
  const dTD=tocDo===0?ts.tocDo:Math.max(0,ts.tocDo-(tocDo/60)*ts.tocDo);
  return json({success:true,data:{tenNV,period:{loai,tuNgay:fmtDate(startDate),denNgay:fmtDate(endDate)},soHoanThanh:ht.length,soTraViec:tv.length,tocDoTB:tocDo,trongSo:ts,tongDiem:Math.min(Math.round(dHT+dTV+dTD),100)}});
}

// ── DASHBOARD (tối ưu: 1 pass dữ liệu) ──
function getDashboard(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const ngay=p.ngay||fmtDate(new Date());
  const data=getSheet(SHEETS.CONG_VIEC).getDataRange().getValues();
  const DS_LOAI=['Ca','Định kỳ','Phát sinh'], DS_CA=['Sáng','Trưa','Chiều','Tối'];
  const eb=()=>({chuaLam:0,dangLam:0,hoanThanh:0,traViec:0});
  const theoLoai={};
  DS_LOAI.forEach(l=>{theoLoai[l]={_all:eb()};DS_CA.forEach(c=>{theoLoai[l][c]=eb();});});
  const stats=eb(),dl=[],cl=[];
  for(let i=1;i<data.length;i++){
    const row=data[i]; if(!row[0]) continue;
    if(readDateFromSheet(row[5])!==ngay) continue;
    const loai=String(row[3]||'').trim(), ca=String(row[4]||'').trim(), tt=String(row[6]||'').trim();
    if(tt===TRANG_THAI.CHUA_LAM)        stats.chuaLam++;
    else if(tt===TRANG_THAI.DANG_LAM)   stats.dangLam++;
    else if(tt===TRANG_THAI.HOAN_THANH) stats.hoanThanh++;
    else if(tt===TRANG_THAI.TRA_VIEC)   stats.traViec++;
    const lk=DS_LOAI.includes(loai)?loai:'Phát sinh';
    const tk=tt===TRANG_THAI.CHUA_LAM?'chuaLam':tt===TRANG_THAI.DANG_LAM?'dangLam':tt===TRANG_THAI.HOAN_THANH?'hoanThanh':tt===TRANG_THAI.TRA_VIEC?'traViec':null;
    if(tk){theoLoai[lk]['_all'][tk]++;if(DS_CA.includes(ca))theoLoai[lk][ca][tk]++;}
    const v=rowToViec(row);
    if(tt===TRANG_THAI.DANG_LAM)dl.push(v); if(tt===TRANG_THAI.CHUA_LAM)cl.push(v);
  }
  const tbData=getSheet(SHEETS.THONG_BAO).getDataRange().getValues();
  const ddSet=getDaDocSet(user.tenNV);
  const tbChuaDoc=tbData.slice(1).filter(r=>r[0]&&(r[3]==='Tất cả'||r[3]===user.tenNV)&&!ddSet.has(String(r[0]))).length;
  return json({success:true,ngay,stats,theoLoai,viecDangLam:dl,viecChuaLam:cl,tbChuaDoc,caHienTai:getCaHienTai()});
}

// ── NHÂN VIÊN ──
function getNhanVien(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const data=getSheet(SHEETS.NHAN_VIEN).getDataRange().getValues(), result=[];
  for(let i=1;i<data.length;i++){
    const r=data[i]; if(!r[0]) continue;
    result.push({id:String(r[0]).trim(),tenNV:String(r[1]).trim(),matKhau:String(r[2]).trim(),
      vaiTro:String(r[3]).trim(),email:String(r[4]||'').trim(),trangThai:String(r[5]).trim()});
  }
  return json({success:true,data:result});
}
function luuNhanVien(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const ten=(p.tenNV||'').trim(), mk=(p.matKhau||'').trim();
  const vt=(p.vaiTro||'Nhân viên').trim(), em=(p.email||'').trim();
  const tt=(p.trangThai||'Hoạt động').trim(), eid=(p.id||'').trim();
  if(!ten) return json({success:false,error:'Thiếu tên đăng nhập'});
  const sh=getSheet(SHEETS.NHAN_VIEN), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    const r=data[i]; if(!r[0]) continue;
    if(String(r[1]).trim().toLowerCase()===ten.toLowerCase()&&String(r[0]).trim()!==eid)
      return json({success:false,error:'Tên "'+ten+'" đã tồn tại'});
  }
  if(eid){
    const ri=findById(sh,eid); if(!ri) return json({success:false,error:'Không tìm thấy'});
    sh.getRange(ri,2).setValue(ten); if(mk)sh.getRange(ri,3).setValue(mk);
    sh.getRange(ri,4).setValue(vt); sh.getRange(ri,5).setValue(em); sh.getRange(ri,6).setValue(tt);
    if(mk||tt==='Nghỉ việc') sh.getRange(ri,7,1,2).setValues([['','']]);
    return json({success:true});
  }else{
    if(!mk||mk.length!==6) return json({success:false,error:'Mật khẩu phải đúng 6 ký tự'});
    let maxNum=0;
    for(let i=1;i<data.length;i++){const id=String(data[i][0]||'');if(id.startsWith('NV')){const n=parseInt(id.replace('NV',''))||0;if(n>maxNum)maxNum=n;}}
    const newId='NV'+String(maxNum+1).padStart(3,'0');
    sh.appendRow([newId,ten,mk,vt,em,tt,'','']);
    return json({success:true,id:newId});
  }
}

// ── BÁO CÁO ──
function getBaoCao(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const start=parseDate(p.tuNgay||fmtDate(new Date())), end=parseDate(p.denNgay||fmtDate(new Date()));
  end.setHours(23,59,59);
  const data=getSheet(SHEETS.CONG_VIEC).getDataRange().getValues();
  let tong=0,ht=0,tv=0,chua=0; const theoLoai={},theoNV={};
  for(let i=1;i<data.length;i++){
    const row=data[i]; if(!row[0]) continue;
    const d=dateFromSheet(row[5]); if(!d||d<start||d>end) continue;
    tong++;
    theoLoai[row[3]]=(theoLoai[row[3]]||0)+1;
    if(row[6]===TRANG_THAI.HOAN_THANH){ht++;const nv=row[7]||'—';if(!theoNV[nv])theoNV[nv]={hoanThanh:0,traViec:0};theoNV[nv].hoanThanh++;}
    if(row[6]===TRANG_THAI.CHUA_LAM||row[6]===TRANG_THAI.DANG_LAM)chua++;
  }
  return json({success:true,data:{tongViec:tong,hoanThanh:ht,traViec:tv,chuaLam:chua,theoLoai,theoNV,tuNgay:p.tuNgay,denNgay:p.denNgay}});
}

// ── CÀI ĐẶT ──
function getCaiDat(p,user){if(user.vaiTro!=='Quản lý')return json({success:false,error:'Không có quyền'});return json({success:true,data:getCaiDatObj()});}
function getCaiDatObj(){
  const data=getSheet(SHEETS.CAI_DAT).getDataRange().getValues(); const obj={};
  for(let i=1;i<data.length;i++){if(data[i][0]){try{obj[data[i][0]]=JSON.parse(data[i][1]);}catch(e){obj[data[i][0]]=data[i][1];}}}
  return obj;
}
function luuCaiDat(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  if(!p.key) return json({success:false,error:'Thiếu key'});
  const sh=getSheet(SHEETS.CAI_DAT), data=sh.getDataRange().getValues();
  const val=typeof p.value==='object'?JSON.stringify(p.value):p.value;
  let found=false;
  for(let i=1;i<data.length;i++){if(data[i][0]===p.key){sh.getRange(i+1,2).setValue(val);found=true;break;}}
  if(!found)sh.appendRow([p.key,val]);
  return json({success:true});
}

// ── ĐĂNG XUẤT ──
function dangXuat(p,user){
  const sh=getSheet(SHEETS.NHAN_VIEN), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][6]||'')===String(p.token||'')){
      sh.getRange(i+1,7,1,2).setValues([['','']]);
      break;
    }
  }
  return json({success:true});
}

// ── PHÂN CÔNG & KPI ──
function ghiPhanCong(cvId,tenViec,nv,hd,tg){getSheet(SHEETS.PHAN_CONG).appendRow(['PC_'+Date.now(),cvId,tenViec,nv,hd,tg.toISOString()]);}
function capNhatKPI(tenNV,tgNhan,tgHT){const phut=Math.round((tgHT-tgNhan)/60000);getSheet(SHEETS.KPI).appendRow(['KPI_'+Date.now(),tenNV,getTuan(tgHT),tgHT.getMonth()+1,tgHT.getFullYear(),1,0,phut,1]);}

// ── TRIGGER ──
function triggerTaoViecDinhKy(){
  const today=new Date(),thu=today.getDay(),ngayThg=today.getDate(),todayStr=fmtDate(today);
  const tmpl=getSheet(SHEETS.VIEC_CA).getDataRange().getValues();
  const cvSh=getSheet(SHEETS.CONG_VIEC), exist=cvSh.getDataRange().getValues().slice(1);
  let soTao=0, soBo=0;
  for(let i=1;i<tmpl.length;i++){
    const row=tmpl[i];
    const loai=String(row[3]||'').trim(); if(!loai){soBo++;continue;}
    const ten=String(row[1]||'').trim(); if(!ten){soBo++;continue;}
    if(String(row[8]||'').trim()==='Tạm dừng') continue;
    const ca=String(row[4]||'').trim(), chu=String(row[5]||'').trim();
    let create=false;
    if(loai===LOAI_VIEC.CA) create=true;
    else if(loai===LOAI_VIEC.DINH_KY){
      if(chu==='Hàng ngày')create=true;
      else if(chu==='Hàng tuần')create=(thu===(parseInt(row[6])||1));
      else if(chu==='Hàng tháng')create=(ngayThg===(parseInt(row[7])||1));
      else create=true;
    }else create=true;
    if(!create) continue;
    const dup=exist.find(r=>String(r[1]).trim()===ten&&String(r[3]).trim()===loai&&readDateFromSheet(r[5])===todayStr);
    if(dup) continue;
    writeViec([`CV_${Date.now()}_${i}`,ten,String(row[2]||''),loai,ca,todayStr,TRANG_THAI.CHUA_LAM,'','','','Bình thường','',chu,todayStr,'Hệ thống']);
    soTao++; Utilities.sleep(30);
  }
  ghiNhatKy('TRIGGER','triggerTaoViecDinhKy','Hệ thống',`${todayStr}: tạo=${soTao} bỏ=${soBo}`);
}

// ── SETUP ──
function setupSheets(){setupNhanVien();setupCaLamViec();setupDanhSachViecCa();setupCongViec();setupPhanCong();setupThongBao();setupKPI();setupCaiDat();setupNhatKy();setupDaDocThongBao();setupTrigger();SpreadsheetApp.getUi().alert('✅ Setup KIMMIZU hoàn tất!');}
function setupNhanVien(){let sh=SS.getSheetByName(SHEETS.NHAN_VIEN);const isNew=!sh;if(isNew)sh=SS.insertSheet(SHEETS.NHAN_VIEN);sh.getRange(1,1,1,8).setValues([['ID','TenNV','MatKhau','VaiTro','Email','TrangThai','Token','TokenExpiry']]);styleHeader(sh,8);if(isNew||sh.getLastRow()<2){sh.getRange(2,1,2,8).setValues([['NV001','Quan ly','admin1','Quản lý','','Hoạt động','',''],['NV002','Nhan vien 1','nv1111','Nhân viên','','Hoạt động','','']]);Logger.log('NhanVien: Điền mẫu');}setDropdown(sh,'D2:D50',['Quản lý','Nhân viên']);setDropdown(sh,'F2:F50',['Hoạt động','Nghỉ việc']);sh.setFrozenRows(1);}
function setupCaLamViec(){let sh=SS.getSheetByName(SHEETS.CA_LAM_VIEC)||SS.insertSheet(SHEETS.CA_LAM_VIEC);sh.getRange(1,1,1,4).setValues([['TenCa','GioBatDau','GioKetThuc','TrangThai']]);styleHeader(sh,4);if(sh.getLastRow()<2)sh.getRange(2,1,4,4).setValues([['Sáng','07:00','10:00','Hoạt động'],['Trưa','10:00','14:00','Hoạt động'],['Chiều','14:00','18:00','Hoạt động'],['Tối','18:00','22:00','Hoạt động']]);sh.setFrozenRows(1);}
function setupDanhSachViecCa(){let sh=SS.getSheetByName(SHEETS.VIEC_CA)||SS.insertSheet(SHEETS.VIEC_CA);sh.getRange(1,1,1,9).setValues([['ID','TenViec','MoTa','LoaiViec','Ca','ChuKy','ThuTrongTuan','NgayTrongThang','TrangThai']]);styleHeader(sh,9);setDropdown(sh,'D2:D50',['Ca','Định kỳ']);setDropdown(sh,'E2:E50',['Sáng','Trưa','Chiều','Tối','']);setDropdown(sh,'F2:F50',['Hàng ngày','Hàng tuần','Hàng tháng','']);setDropdown(sh,'I2:I50',['Hoạt động','Tạm dừng']);sh.setFrozenRows(1);}
function setupCongViec(){let sh=SS.getSheetByName(SHEETS.CONG_VIEC)||SS.insertSheet(SHEETS.CONG_VIEC);sh.getRange(1,1,1,15).setValues([['ID','TenViec','MoTa','LoaiViec','Ca','NgayThucHien','TrangThai','NguoiNhan','ThoiGianNhan','ThoiGianHoanThanh','UuTien','HanChot','ChuKy','NgayTao','NguoiTao']]);styleHeader(sh,15);setDropdown(sh,'D2:D500',['Ca','Định kỳ','Phát sinh']);setDropdown(sh,'E2:E500',['Sáng','Trưa','Chiều','Tối','']);setDropdown(sh,'G2:G500',['Chưa làm','Đang làm','Hoàn thành','Trả việc']);sh.getRange(2,6,500,1).setNumberFormat('@');sh.getRange(2,12,500,1).setNumberFormat('@');sh.getRange(2,14,500,1).setNumberFormat('@');sh.setFrozenRows(1);}
function setupPhanCong(){let sh=SS.getSheetByName(SHEETS.PHAN_CONG)||SS.insertSheet(SHEETS.PHAN_CONG);sh.getRange(1,1,1,6).setValues([['ID','CongViecID','TenViec','NhanVien','HanhDong','ThoiGian']]);styleHeader(sh,6);sh.setFrozenRows(1);}
function setupThongBao(){let sh=SS.getSheetByName(SHEETS.THONG_BAO)||SS.insertSheet(SHEETS.THONG_BAO);sh.getRange(1,1,1,7).setValues([['ID','TieuDe','NoiDung','NguoiNhan','NgayTao','DaDoc','NguoiTao']]);styleHeader(sh,7);sh.setFrozenRows(1);}
function setupKPI(){let sh=SS.getSheetByName(SHEETS.KPI)||SS.insertSheet(SHEETS.KPI);sh.getRange(1,1,1,9).setValues([['ID','NhanVien','Tuan','Thang','Nam','SoHoanThanh','SoTraViec','TongPhut','SoLanTinhTD']]);styleHeader(sh,9);sh.setFrozenRows(1);}
function setupCaiDat(){let sh=SS.getSheetByName(SHEETS.CAI_DAT)||SS.insertSheet(SHEETS.CAI_DAT);sh.getRange(1,1,1,3).setValues([['Key','Value','MoTa']]);styleHeader(sh,3);if(sh.getLastRow()<2)sh.getRange(2,1,4,3).setValues([['caLamViec',JSON.stringify([{ten:'Sáng',batDau:'07:00',ketThuc:'10:00'},{ten:'Trưa',batDau:'10:00',ketThuc:'14:00'},{ten:'Chiều',batDau:'14:00',ketThuc:'18:00'},{ten:'Tối',batDau:'18:00',ketThuc:'22:00'}]),'Cấu hình ca'],['kpiTrongSo',JSON.stringify({hoanThanh:60,traViec:25,tocDo:15}),'Trọng số KPI'],['tenQuan','KIMMIZU','Tên quán'],['phienBan','3.0','Phiên bản']]);sh.setFrozenRows(1);}
function setupDaDocThongBao(){let sh=SS.getSheetByName(SHEETS.DA_DOC)||SS.insertSheet(SHEETS.DA_DOC);sh.getRange(1,1,1,3).setValues([['ThongBaoID','TenNV','ThoiGian']]);styleHeader(sh,3);sh.setFrozenRows(1);}
function setupNhatKy(){let sh=SS.getSheetByName(SHEETS.NHAT_KY)||SS.insertSheet(SHEETS.NHAT_KY);sh.getRange(1,1,1,5).setValues([['ThoiGian','Loai','Action','NguoiDung','GhiChu']]);styleHeader(sh,5);sh.setFrozenRows(1);}
function setupTrigger(){
  ScriptApp.getProjectTriggers().forEach(t=>{if(['triggerTaoViecDinhKy','onEditSheet'].includes(t.getHandlerFunction()))ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('triggerTaoViecDinhKy').timeBased().everyDays(1).atHour(0).create();
  ScriptApp.newTrigger('onEditSheet').forSpreadsheet(SS).onEdit().create();
}

// ── HELPERS (tối ưu) ──
const _sheetCache={};
function getSheet(name){if(_sheetCache[name])return _sheetCache[name];const s=SS.getSheetByName(name);if(!s)throw new Error('Sheet "'+name+'" không tồn tại');_sheetCache[name]=s;return s;}
function findById(sh,id){const d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++){if(String(d[i][0])===String(id))return i+1;}return null;}
function writeViec(rowData){const sh=getSheet(SHEETS.CONG_VIEC),r=sh.getLastRow()+1;sh.getRange(r,1,1,15).setValues([rowData]);sh.getRange(r,6).setNumberFormat('@');sh.getRange(r,12).setNumberFormat('@');sh.getRange(r,14).setNumberFormat('@');}
function readDateFromSheet(val){if(!val&&val!==0)return'';if(val instanceof Date){if(isNaN(val))return'';return fmtDate(val);}const s=String(val).trim();if(!s)return'';if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s))return s;if(/^\d{1,2}-\d{1,2}-\d{4}$/.test(s)){const p=s.split('-');return pad(p[0])+'/'+pad(p[1])+'/'+p[2];}if(/^\d{4}-\d{2}-\d{2}/.test(s)){const p=s.substring(0,10).split('-');return pad(p[2])+'/'+pad(p[1])+'/'+p[0];}const d=new Date(val);return isNaN(d)?'':fmtDate(d);}
function dateFromSheet(val){if(!val)return null;if(val instanceof Date)return isNaN(val)?null:val;const s=String(val).trim();if(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(s)){const p=s.split(/[\/\-]/);return new Date(parseInt(p[2]),parseInt(p[1])-1,parseInt(p[0]));}if(/^\d{4}-\d{2}-\d{2}/.test(s)){const p=s.substring(0,10).split('-');return new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]));}const d=new Date(val);return isNaN(d)?null:d;}
function parseDate(s){if(!s)return new Date();if(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(s)){const p=s.split(/[\/\-]/);return new Date(parseInt(p[2]),parseInt(p[1])-1,parseInt(p[0]));}return new Date(s);}
function fmtDate(d){if(!d||isNaN(d))return'';return pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear();}
function fmtDateTime(d){if(!d||isNaN(d))return'';return fmtDate(d)+' '+pad(d.getHours())+':'+pad(d.getMinutes());}
function pad(n){return String(n).padStart(2,'0');}
function getTuan(d){const s=new Date(d.getFullYear(),0,1);return Math.ceil(((d-s)/86400000+s.getDay()+1)/7);}
function json(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
function styleHeader(sh,n){sh.getRange(1,1,1,n).setBackground('#3B2314').setFontColor('#F5E6C8').setFontWeight('bold').setHorizontalAlignment('center');}
function setDropdown(sh,range,values){sh.getRange(range).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(values).build());}
function getLoaiLog(action){const m={'nhanViec':'NHẬN VIỆC','hoanThanhViec':'HOÀN THÀNH','traViec':'TRẢ VIỆC','taoViecPhatSinh':'TẠO PHÁT SINH','taoCongViec':'TẠO CÔNG VIỆC','suaCongViec':'SỬA CÔNG VIỆC','xoaCongViec':'XOÁ CÔNG VIỆC','taoThongBao':'TẠO THÔNG BÁO','suaThongBao':'SỬA THÔNG BÁO','xoaThongBao':'XOÁ THÔNG BÁO','docThongBao':'ĐỌC THÔNG BÁO','luuNhanVien':'LƯU NHÂN VIÊN','luuCaiDat':'LƯU CÀI ĐẶT','nhacNhoThongBao':'NHẮC EMAIL','login':'ĐĂNG NHẬP','dangXuat':'ĐĂNG XUẤT'};return m[action]||action;}

// ── onEdit TRIGGER ──
function onEditSheet(e){try{const range=e.range,sh=range.getSheet(),shName=sh.getName(),row=range.getRow(),col=range.getColumn();if(shName===SHEETS.CONG_VIEC)handleEditCongViec(sh,row,col,e);if(shName===SHEETS.NHAN_VIEN)handleEditNhanVien(sh,row,col,e);if(shName===SHEETS.VIEC_CA)handleEditViecCa(sh,row,col,e);if(shName===SHEETS.CAI_DAT)handleEditCaiDat(sh,row,col,e);}catch(err){Logger.log('onEditSheet: '+err);}}
function handleEditCongViec(sh,row,col,e){if(row<2)return;const rd=sh.getRange(row,1,1,15).getValues()[0];const ten=String(rd[1]||'').trim(),ca=String(rd[4]||'').trim();if(ten&&!rd[0]){const id='CV_'+Date.now(),loai=String(rd[3]||LOAI_VIEC.PHAT_SINH).trim();const ngay=rd[5]?readDateFromSheet(rd[5]):fmtDate(new Date());sh.getRange(row,6).setNumberFormat('@');sh.getRange(row,12).setNumberFormat('@');sh.getRange(row,14).setNumberFormat('@');sh.getRange(row,1).setValue(id);if(!rd[3])sh.getRange(row,4).setValue(LOAI_VIEC.PHAT_SINH);if(loai===LOAI_VIEC.CA&&!ca)sh.getRange(row,5).setValue('Sáng');sh.getRange(row,6).setValue(ngay);if(!rd[6])sh.getRange(row,7).setValue(TRANG_THAI.CHUA_LAM);if(!rd[10])sh.getRange(row,11).setValue('Bình thường');sh.getRange(row,14).setValue(fmtDate(new Date()));if(!rd[14])sh.getRange(row,15).setValue('Sheet');return;}if(rd[0]&&(col===6||col===12||col===14)){const cell=sh.getRange(row,col);cell.setNumberFormat('@');const v=readDateFromSheet(cell.getValue());if(v)cell.setValue(v);}}
function handleEditNhanVien(sh,row,col,e){if(row<2)return;const rd=sh.getRange(row,1,1,8).getValues()[0];const ten=String(rd[1]||'').trim();if(ten&&!rd[0]){const all=sh.getDataRange().getValues();let maxNum=0;for(let i=1;i<all.length;i++){const n=parseInt(String(all[i][0]||'').replace('NV',''))||0;if(n>maxNum)maxNum=n;}const newId='NV'+String(maxNum+1).padStart(3,'0');sh.getRange(row,1).setValue(newId);if(!rd[3])sh.getRange(row,4).setValue('Nhân viên');if(!rd[5])sh.getRange(row,6).setValue('Hoạt động');}}
function handleEditViecCa(sh,row,col,e){if(row<2)return;const rd=sh.getRange(row,1,1,9).getValues()[0];if(!String(rd[0]||'').trim()&&!String(rd[1]||'').trim())return;}
function handleEditCaiDat(sh,row,col,e){if(row<2||col!==2)return;}

// ── XUẤT EXCEL ──
function xuatBaoCaoExcel(p,user){
  if(user.vaiTro!=='Quản lý') return json({success:false,error:'Không có quyền'});
  const tuNgay   = p.tuNgay   || fmtDate(new Date());
  const denNgay  = p.denNgay  || fmtDate(new Date());
  const loai     = p.loai     || ''; // '' = tất cả
  const start    = parseDate(tuNgay), end = parseDate(denNgay);
  end.setHours(23,59,59);

  // Lọc CongViec theo khoảng ngày và loại
  const cvData   = getSheet(SHEETS.CONG_VIEC).getDataRange().getValues();
  const nvData   = getSheet(SHEETS.NHAN_VIEN).getDataRange().getValues();
  const pcData   = getSheet(SHEETS.PHAN_CONG).getDataRange().getValues();

  // Tạo Spreadsheet mới để export
  const ssName   = `KIMMIZU_BaoCao_${tuNgay.replace(/\//g,'-')}_${denNgay.replace(/\//g,'-')}`;
  const newSS    = SpreadsheetApp.create(ssName);

  // ── Sheet 1: Tổng hợp công việc ──
  const sh1 = newSS.getActiveSheet();
  sh1.setName('Công Việc');

  // Header
  const h1 = ['ID','Tên Việc','Loại','Ca','Ngày','Trạng Thái','Người Nhận','Giờ Nhận','Giờ HT','Ưu Tiên','Người Tạo'];
  sh1.getRange(1,1,1,h1.length).setValues([h1]).setBackground('#3B2314').setFontColor('#F5E6C8').setFontWeight('bold');

  const rows = [];
  for(let i=1;i<cvData.length;i++){
    const row=cvData[i]; if(!row[0]) continue;
    const d=dateFromSheet(row[5]); if(!d||d<start||d>end) continue;
    if(loai && String(row[3]).trim()!==loai) continue;
    rows.push([
      row[0], row[1], row[3], row[4],
      readDateFromSheet(row[5]), row[6], row[7]||'',
      row[8]?fmtDateTime(new Date(row[8])):'',
      row[9]?fmtDateTime(new Date(row[9])):'',
      row[10]||'', row[14]||''
    ]);
  }
  if(rows.length) sh1.getRange(2,1,rows.length,h1.length).setValues(rows);

  // Màu theo trạng thái
  for(let i=0;i<rows.length;i++){
    const tt=rows[i][5];
    const color=tt==='Hoàn thành'?'#E8F5EE':tt==='Đang làm'?'#E8F0FE':tt==='Trả việc'?'#FDECEA':'#FEF0E6';
    sh1.getRange(i+2,1,1,h1.length).setBackground(color);
  }
  sh1.autoResizeColumns(1,h1.length);
  sh1.setFrozenRows(1);

  // ── Sheet 2: Thống kê nhân viên ──
  const sh2 = newSS.insertSheet('Thống Kê NV');
  const h2  = ['Nhân Viên','Hoàn Thành','Trả Việc','Tổng Thao Tác'];
  sh2.getRange(1,1,1,h2.length).setValues([h2]).setBackground('#3B2314').setFontColor('#F5E6C8').setFontWeight('bold');

  const nvStats={};
  for(let i=1;i<pcData.length;i++){
    const r=pcData[i]; if(!r[0]) continue;
    const t=new Date(r[5]); if(t<start||t>end) continue;
    const nv=String(r[3]||''); if(!nv) continue;
    if(!nvStats[nv]) nvStats[nv]={ht:0,tv:0};
    if(r[4]==='Hoàn thành') nvStats[nv].ht++;
    if(String(r[4]).startsWith('Trả việc')) nvStats[nv].tv++;
  }
  const nvRows = Object.entries(nvStats).map(([nv,s])=>[nv,s.ht,s.tv,s.ht+s.tv]);
  nvRows.sort((a,b)=>b[1]-a[1]); // sort theo hoàn thành
  if(nvRows.length) sh2.getRange(2,1,nvRows.length,h2.length).setValues(nvRows);
  sh2.autoResizeColumns(1,h2.length);
  sh2.setFrozenRows(1);

  // ── Sheet 3: Thống kê theo ngày ──
  const sh3 = newSS.insertSheet('Thống Kê Ngày');
  const h3  = ['Ngày','Ca','Tổng Việc','Hoàn Thành','Chưa/Đang Làm','Trả Việc','Tỉ Lệ HT %'];
  sh3.getRange(1,1,1,h3.length).setValues([h3]).setBackground('#3B2314').setFontColor('#F5E6C8').setFontWeight('bold');

  const ngayStats={};
  for(let i=1;i<cvData.length;i++){
    const row=cvData[i]; if(!row[0]) continue;
    const ngay=readDateFromSheet(row[5]); if(!ngay) continue;
    const d=dateFromSheet(row[5]); if(!d||d<start||d>end) continue;
    const ca=String(row[4]||'(Khác)'), key=ngay+'__'+ca;
    if(!ngayStats[key]) ngayStats[key]={ngay,ca,tong:0,ht:0,chua:0,tv:0};
    ngayStats[key].tong++;
    const tt=String(row[6]||'');
    if(tt==='Hoàn thành') ngayStats[key].ht++;
    else if(tt==='Trả việc') ngayStats[key].tv++;
    else ngayStats[key].chua++;
  }
  const ngayRows = Object.values(ngayStats).map(s=>[
    s.ngay, s.ca, s.tong, s.ht, s.chua, s.tv,
    s.tong>0?Math.round(s.ht/s.tong*100)+'%':'0%'
  ]);
  ngayRows.sort((a,b)=>a[0].localeCompare(b[0]));
  if(ngayRows.length) sh3.getRange(2,1,ngayRows.length,h3.length).setValues(ngayRows);
  sh3.autoResizeColumns(1,h3.length);
  sh3.setFrozenRows(1);

  // Tạo link download Excel
  const fileId = newSS.getId();
  const exportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx&exportFormat=xlsx`;

  ghiNhatKy('XUẤT EXCEL','xuatBaoCaoExcel',user.tenNV,
    `${tuNgay}→${denNgay} loai=${loai||'tất cả'} ${rows.length} dòng`);

  return json({success:true, fileId, exportUrl,
    fileName: ssName+'.xlsx',
    tongViec: rows.length,
    message: 'Đã tạo file Excel. Click link để tải.'
  });
}

// ── TIỆN ÍCH ──
function taoViecHomNay(){triggerTaoViecDinhKy();SpreadsheetApp.getUi().alert('Đã tạo việc hôm nay.');}
function fixCaToanBo(){const tmpl=getSheet(SHEETS.VIEC_CA).getDataRange().getValues();const caMap={};for(let i=1;i<tmpl.length;i++){const ten=String(tmpl[i][1]||'').trim(),loai=String(tmpl[i][3]||'').trim(),ca=String(tmpl[i][4]||'').trim();if(ten&&loai==='Ca'&&ca)caMap[ten]=ca;}if(!Object.keys(caMap).length){SpreadsheetApp.getUi().alert('DanhSachViecCa chưa có Ca.');return;}const cvSh=getSheet(SHEETS.CONG_VIEC),cvData=cvSh.getDataRange().getValues();let fixed=0;for(let i=1;i<cvData.length;i++){const row=cvData[i];if(!row[0]||String(row[3]).trim()!=='Ca')continue;const caDung=caMap[String(row[1]).trim()];if(caDung&&String(row[4]).trim()!==caDung){cvSh.getRange(i+1,5).setValue(caDung);fixed++;}}SpreadsheetApp.getUi().alert('✅ Đã fix '+fixed+' dòng.');}
function testNhatKy(){ghiNhatKy('TEST','testNhatKy','Admin','Test lúc '+fmtDateTime(new Date()));SpreadsheetApp.getUi().alert('✅ OK - kiểm tra tab NhatKy');}
