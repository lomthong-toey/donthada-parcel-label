const BRANCHES=['งามวงศ์วาน','บางพลี','บางกะปิ','เยาวราช','จตุจักร (JJ Mall)','ปากเกร็ด','นวมินทร์','เมืองทองธานี','รามอินทรา'];
const OPTIONS={visit:['ครั้งแรก','เคยมาแล้ว'],metal:['ทอง / นาก','เงิน','ทั้งสองอย่าง','สอบถามข้อมูล'],source:['ของขวัญ / มรดก','ซื้อไว้ใช้เอง','สะสม / ลงทุน','ธุรกิจ / ร้านค้า','อื่น ๆ','ไม่สะดวกตอบ'],purpose:['เปลี่ยนเป็นเงิน','ตรวจมูลค่า / เปอร์เซ็นต์','ราคาน่าสนใจ','ของไม่ได้ใช้','หมุนเวียนธุรกิจ','สำรวจราคา'],priority:['ราคา','ความน่าเชื่อถือ','ความโปร่งใส','ความรวดเร็ว','ความสะดวก','การดูแล'],melt:['หลอมตรวจเปอร์เซ็นต์','ขายเป็นชิ้นงาน','ขอข้อมูลก่อน','ได้ทั้งสองแบบ'],price:['ราคาสุทธิ','สูตรตามเปอร์เซ็นต์','ราคาต่อกรัม','ให้พนักงานอธิบาย'],status:['ขายแล้ว','ยังไม่ตัดสินใจ','ยังไม่ขาย','กำลังรอผล'],reason:['ราคาไม่ตรงที่คาด','กังวลเรื่องหลอม','อยากได้ราคาที่แน่นอน','เทียบร้านอื่น','ปรึกษาครอบครัว','ยังไม่พร้อม','อื่น ๆ'],channel:['Google','Facebook','TikTok','LINE','YouTube','คนรู้จัก','ป้ายหน้าร้าน','อื่น ๆ'],wish:['แจ้งเตือนราคา','เช็กสถานะคิว','ประเมินล่วงหน้า','บริการเร็วขึ้น','สิทธิพิเศษ','คำแนะนำ']};
function json(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){return json({ok:true,service:'Pantip Customer Insight'});}
function doPost(e){
let lock;
try {
const raw=e.postData.contents;
if(raw.length>4096)throw Error('Invalid payload');
const p=JSON.parse(raw),a=p.answers;
if(!/^B0[1-9]$/.test(p.branch)||!a||! /^[a-zA-Z0-9-]{16,64}$/.test(p.id))throw Error('Invalid input');
const vals=Object.keys(OPTIONS).map(k=>{if(k==='reason'&&![1,2].includes(a.status))return '';if(!Number.isInteger(a[k])||a[k]<0||a[k]>=OPTIONS[k].length)throw Error('Invalid answer');return OPTIONS[k][a[k]];});
lock=LockService.getScriptLock();lock.waitLock(10000);
const sheet=SpreadsheetApp.openById('1-g4nYg5ORo4Q2kVL2pd__n_8S7qxmjLos8IhINiduw8').getSheetByName('Responses');
if(sheet.getLastRow()>1&&sheet.getRange(2,2,sheet.getLastRow()-1,1).createTextFinder(p.id).matchEntireCell(true).findNext())return json({ok:true,duplicate:true});
sheet.appendRow([new Date(),p.id,p.branch,BRANCHES[Number(p.branch.slice(1))-1],...vals]);return json({ok:true});
}catch(err){return json({ok:false,error:'Unable to save response'});}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}