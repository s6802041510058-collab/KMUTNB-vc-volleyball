/** KMUTNB VC — Google Apps Script backend. See SETUP.md. */
const TABLES=['members','posts','hearts','events','joins','equipment','loans','folders','files','sessions','otps','settings'];
function prop(k){return PropertiesService.getScriptProperties().getProperty(k)||''}
function db(){if(!prop('SPREADSHEET_ID'))throw Error('ยังไม่ได้ตั้งค่า SPREADSHEET_ID');return SpreadsheetApp.openById(prop('SPREADSHEET_ID'))}
function sheet(t){if(!TABLES.includes(t))throw Error('Invalid table');let s=db().getSheetByName(t);if(!s){s=db().insertSheet(t);s.appendRow(['id','json'])}return s}
function rows(t){let s=sheet(t);if(s.getLastRow()<2)return [];return s.getRange(2,2,s.getLastRow()-1,1).getValues().map(r=>JSON.parse(r[0]))}
function put(t,o){let s=sheet(t),ids=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,1).getValues().flat():[],i=ids.indexOf(o.id);if(i>=0)s.getRange(i+2,1,1,2).setValues([[o.id,JSON.stringify(o)]]);else s.appendRow([o.id,JSON.stringify(o)]);return o}
function del(t,id){let s=sheet(t);if(s.getLastRow()<2)return;let ids=s.getRange(2,1,s.getLastRow()-1,1).getValues().flat(),i=ids.indexOf(id);if(i>=0)s.deleteRow(i+2)}
function uid(){return Utilities.getUuid()}
function now(){return new Date().toISOString()}
function hash(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s).map(x=>('0'+((x+256)%256).toString(16)).slice(-2)).join('')}
function text(s,max){return String(s||'').trim().slice(0,max||5000)}
function publicMember(m){return {id:m.id,name:m.name,position:m.position}}
function current(token){if(!token)return null;let s=rows('sessions').find(x=>x.id===hash(token)&&x.expires>Date.now());return s?rows('members').find(m=>m.id===s.memberId):null}
function must(m){if(!m)throw Error('กรุณาเข้าสู่ระบบอีกครั้ง');return m}
function admin(m){must(m);if(m.email!==prop('ADMIN_EMAIL').toLowerCase().trim())throw Error('เฉพาะผู้ดูแลเท่านั้น')}
function folder(){if(!prop('DRIVE_FOLDER_ID'))throw Error('ยังไม่ได้ตั้งค่า DRIVE_FOLDER_ID');return DriveApp.getFolderById(prop('DRIVE_FOLDER_ID'))}
function upload(f,imageOnly){if(!f||!f.base64||!f.name)throw Error('กรุณาแนบไฟล์');if(imageOnly&&!/^image\/(jpeg|png|webp)$/.test(f.mime))throw Error('ต้องเป็นภาพ JPEG PNG หรือ WebP');let b=Utilities.base64Decode(f.base64);if(b.length>5*1024*1024)throw Error('ไฟล์เกิน 5 MB');let file=folder().createFile(Utilities.newBlob(b,f.mime,text(f.name,200)));return file.getId()}
function blob(id){let f=DriveApp.getFileById(id),b=f.getBlob();return {name:f.getName(),mime:b.getContentType(),base64:Utilities.base64Encode(b.getBytes())}}
function setup(){TABLES.forEach(sheet);return 'Ready'}
function doPost(e){let lock=LockService.getScriptLock();try{lock.waitLock(30000);let a=JSON.parse(e.postData.contents),m=current(a.token),result=handle(a,m);return json({ok:true,data:result})}catch(e){return json({ok:false,error:e.message})}finally{if(lock.hasLock())lock.releaseLock()}}
function doGet(){return json({ok:true,data:{service:'KMUTNB VC',status:'ready'}})}
function json(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function handle(a,m){
if(a.action==='requestOtp'){
 let email=text(a.email,254).toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw Error('อีเมลไม่ถูกต้อง');let member=rows('members').find(x=>x.email===email);if(!member&&!a.register)throw Error('ยังไม่มีสมาชิกอีเมลนี้ กรุณาสมัครก่อน');if(!member&&!text(a.name,100))throw Error('กรุณาระบุชื่อ');let old=rows('otps').find(x=>x.id===email);if(old&&Date.now()-old.sentAt<60000)throw Error('กรุณารอ 60 วินาทีก่อนขอรหัสอีกครั้ง');let code=String(Math.floor(100000+Math.random()*900000)),salt=uid();put('otps',{id:email,digest:hash(salt+code),salt,expires:Date.now()+600000,sentAt:Date.now(),attempts:0,name:text(a.name,100),studentId:text(a.studentId,30)});MailApp.sendEmail(email,'รหัสเข้าสู่ระบบ KMUTNB VC','รหัส OTP ของคุณคือ '+code+'\nใช้ได้ภายใน 10 นาที');return {sent:true}
}
if(a.action==='verifyOtp'){
 let email=text(a.email,254).toLowerCase(),o=rows('otps').find(x=>x.id===email);if(!o||o.expires<Date.now()||o.attempts>=5)throw Error('รหัสหมดอายุ กรุณาขอรหัสใหม่');o.attempts++;put('otps',o);if(hash(o.salt+text(a.otp,6))!==o.digest)throw Error('รหัสไม่ถูกต้อง');let member=rows('members').find(x=>x.email===email);if(!member)member=put('members',{id:uid(),email,name:o.name,studentId:o.studentId,position:'สมาชิก',createdAt:now()});del('otps',email);let token=uid()+uid();put('sessions',{id:hash(token),memberId:member.id,expires:Date.now()+86400000});return {token,user:{...publicMember(member),email:member.email,role:member.email===prop('ADMIN_EMAIL').toLowerCase().trim()?'admin':'member'}}
}
if(a.action==='load'){
 let isAdmin=m&&m.email===prop('ADMIN_EMAIL').toLowerCase().trim(),hearts=rows('hearts'),joins=rows('joins'),loans=rows('loans');return {apiVersion:2,currentUser:m?{...publicMember(m),email:m.email,role:isAdmin?'admin':'member'}:null,posts:rows('posts').map(p=>({...p,hearts:hearts.filter(h=>h.postId===p.id).length})),members:rows('members').map(publicMember),events:rows('events').map(e=>({...e,count:joins.filter(j=>j.eventId===e.id).length,joined:!!m&&joins.some(j=>j.eventId===e.id&&j.memberId===m.id)})),equipment:rows('equipment').map(e=>({...e,available:!loans.some(l=>l.equipmentId===e.id&&!l.returnedAt)})),loans:m?loans.filter(l=>isAdmin||l.memberId===m.id).map(l=>({id:l.id,equipmentId:l.equipmentId,equipmentName:l.equipmentName,memberName:l.memberName,borrowedAt:l.borrowedAt,returnedAt:l.returnedAt})):[],folders:rows('folders'),files:rows('files').map(f=>({id:f.id,name:f.name,folderId:f.folderId,createdAt:f.createdAt})),settings:rows('settings').find(s=>s.id==='site')||{}}
}
if(a.action==='postImage'){let p=rows('posts').find(p=>p.id===a.id);if(!p||!p.imageDriveId)throw Error('ไม่พบภาพโพสต์');return blob(p.imageDriveId)}
if(a.action==='heart'){
 must(m);if(!rows('posts').some(p=>p.id===a.id))throw Error('ไม่พบโพสต์');let id=m.id+':'+a.id,existing=rows('hearts').find(h=>h.id===id);if(existing)del('hearts',id);else put('hearts',{id,memberId:m.id,postId:a.id});return true
}
if(a.action==='join'){
 must(m);let e=rows('events').find(x=>x.id===a.id);if(!e||!e.open)throw Error('กิจกรรมปิดรับแล้ว');let id=m.id+':'+e.id,joins=rows('joins');if(joins.some(j=>j.id===id)){del('joins',id);return true}if(e.capacity&&joins.filter(j=>j.eventId===e.id).length>=e.capacity)throw Error('กิจกรรมเต็มแล้ว');put('joins',{id,eventId:e.id,memberId:m.id,createdAt:now()});return true
}
if(a.action==='borrow'||a.action==='return'){
 must(m);if(!a.faceConfirmed)throw Error('ต้องยืนยันว่าภาพเห็นใบหน้าและอุปกรณ์');let equipment=rows('equipment').find(x=>x.id===a.id);if(!equipment)throw Error('ไม่พบอุปกรณ์');let loans=rows('loans');if(a.action==='borrow'){if(loans.some(l=>l.equipmentId===a.id&&!l.returnedAt))throw Error('อุปกรณ์ถูกยืมแล้ว');let photo=upload(a.photo,true);put('loans',{id:uid(),equipmentId:a.id,equipmentName:equipment.name,memberId:m.id,memberName:m.name,borrowedAt:now(),borrowPhoto:photo,returnedAt:null})}else{let l=loans.find(x=>x.id===a.loanId&&x.equipmentId===a.id);if(!l||l.memberId!==m.id||l.returnedAt)throw Error('ไม่พบรายการยืมของคุณ');let photo=upload(a.photo,true);l.returnPhoto=photo;l.returnedAt=now();put('loans',l)}return true
}
if(a.action==='download'){must(m);let f=rows('files').find(x=>x.id===a.id);if(!f)throw Error('ไม่พบไฟล์');return blob(f.driveId)}
if(a.action==='loanPhoto'){must(m);let l=rows('loans').find(x=>x.id===a.id);if(!l)throw Error('ไม่พบรายการ');if(l.memberId!==m.id)admin(m);let id=a.phase==='return'?l.returnPhoto:l.borrowPhoto;if(!id)throw Error('ยังไม่มีภาพ');return blob(id)}
admin(m);
if(a.action==='save'){
 let id=a.id||uid();if(a.kind==='post'){let type=['news','club','university'].includes(a.type)?a.type:'news',old=rows('posts').find(p=>p.id===id);if(!text(a.title)||!text(a.body))throw Error('ระบุหัวข้อและเนื้อหา');let image=text(a.image,2000);if(image&&!/^https:\/\//.test(image))throw Error('URL ภาพต้องเป็น HTTPS');let imageDriveId=a.photo?upload(a.photo,true):(image?'':old?.imageDriveId||'');put('posts',{id,type,imageDriveId,title:text(a.title,200),body:text(a.body,10000),category:text(a.category,100),image,createdAt:old?.createdAt||now()})}
 else if(a.kind==='event'){if(!text(a.title)||!a.date||!text(a.location))throw Error('ข้อมูลกิจกรรมไม่ครบ');let date=new Date(a.date);if(isNaN(date))throw Error('วันที่ไม่ถูกต้อง');put('events',{id,title:text(a.title,200),body:text(a.body),date:date.toISOString(),location:text(a.location,200),capacity:Math.max(0,Math.floor(Number(a.capacity)||0)),open:true})}
 else if(a.kind==='equipment'){if(!text(a.name))throw Error('ระบุชื่ออุปกรณ์');put('equipment',{id,name:text(a.name,150),detail:text(a.detail)})}
 else if(a.kind==='folder'){if(!text(a.name))throw Error('ระบุชื่อโฟลเดอร์');put('folders',{id,name:text(a.name,150)})}
 else if(a.kind==='file'){if(a.folderId&&!rows('folders').some(f=>f.id===a.folderId))throw Error('ไม่พบโฟลเดอร์');let driveId=upload(a.file,false);put('files',{id,name:text(a.file.name,200),driveId,folderId:a.folderId||'',createdAt:now()})}else throw Error('ประเภทไม่ถูกต้อง');return true
}
if(a.action==='delete'){
 if(!['posts','events','equipment','folders','files'].includes(a.table))throw Error('ไม่อนุญาต');if(a.table==='equipment'&&rows('loans').some(l=>l.equipmentId===a.id))throw Error('มีประวัติยืม ไม่สามารถลบได้');if(a.table==='events'&&rows('joins').some(j=>j.eventId===a.id))throw Error('มีผู้ลงชื่อ ไม่สามารถลบได้');if(a.table==='folders'&&rows('files').some(f=>f.folderId===a.id))throw Error('โฟลเดอร์ยังมีไฟล์');if(a.table==='files'){let f=rows('files').find(f=>f.id===a.id);if(f)DriveApp.getFileById(f.driveId).setTrashed(true)}del(a.table,a.id);if(a.table==='posts')rows('hearts').filter(h=>h.postId===a.id).forEach(h=>del('hearts',h.id));return true
}
if(a.action==='settings'){let url=text(a.bookingUrl,2000);if(url&&!/^https:\/\//.test(url))throw Error('URL ต้องเป็น HTTPS');put('settings',{id:'site',bookingUrl:url});return true}
if(a.action==='toggleEvent'){let e=rows('events').find(x=>x.id===a.id);if(!e)throw Error('ไม่พบกิจกรรม');e.open=!e.open;put('events',e);return true}
if(a.action==='position'){let member=rows('members').find(x=>x.id===a.id);if(!member)throw Error('ไม่พบสมาชิก');member.position=text(a.position,100)||'สมาชิก';put('members',member);return true}
if(a.action==='participants'){let ids=rows('joins').filter(j=>j.eventId===a.id).map(j=>j.memberId);return rows('members').filter(m=>ids.includes(m.id)).map(m=>({name:m.name,studentId:m.studentId}))}
throw Error('Unknown action');
}
