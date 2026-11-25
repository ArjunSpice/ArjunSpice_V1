/** Arjun Spice — COD orders, live prices, itemized Sheets, email, cancellation.
 * Script Properties: SPREADSHEET_ID and OWNER_EMAIL (your email address).
 * Set both in Project Settings > Script properties. Deploy as: Execute as Me.
 */
const ORDER_SHEET='Orders', CATALOG_SHEET='Catalog', SETTINGS_SHEET='Settings', ITEM_SHEET='Order Items';
const BASE_HEADERS=['Order ID','Received At','Status','Full Name','Business','Phone','Email','Customer Type','Products (JSON)','Total Quantity (kg)','Address','City','State','PIN','Preferred Delivery Date','Packing','Payment','Notes','Consent','COD Subtotal (INR)','Pricing Status'];
const EXTRA_HEADERS=['Products (Readable)','Product 1','Quantity 1 (kg)','Price 1 (INR/kg)','Amount 1 (INR)','Product 2','Quantity 2 (kg)','Price 2 (INR/kg)','Amount 2 (INR)','Delivery Estimate','Email Notification Status','Cancelled At','Full Delivery Address','Confirmed At','Customer Confirmation Email Status','Admin Action'];
const DELIVERED_SHEET='Delivered Orders';
const ITEM_HEADERS=['Order ID','Received At','Customer Name','Phone','Product','Quantity (kg)','Price per kg (INR)','Line Total (INR)','Order Status'];
function book_(){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(!id)throw Error('Missing SPREADSHEET_ID script property');return SpreadsheetApp.openById(id.trim());}
function ensure_(book,name,headers){let s=book.getSheetByName(name)||book.insertSheet(name);if(s.getLastRow()===0){s.getRange(1,1,1,headers.length).setValues([headers]);s.setFrozenRows(1)}return s;}
function columns_(sheet,headers){const count=sheet.getLastColumn();const existing=count?sheet.getRange(1,1,1,count).getValues()[0].map(String):[];const missing=headers.filter(h=>!existing.includes(h));if(missing.length)sheet.getRange(1,count+1,1,missing.length).setValues([missing]);return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);}
function setupArjunSpice(){const b=book_();columns_(ensure_(b,ORDER_SHEET,BASE_HEADERS),BASE_HEADERS.concat(EXTRA_HEADERS));ensure_(b,ITEM_SHEET,ITEM_HEADERS);const c=ensure_(b,CATALOG_SHEET,['Product','Price per kg (INR)','Active (TRUE/FALSE)']);if(c.getLastRow()===1){c.getRange(2,1,3,3).setValues([['Red Chilli Powder — 341 Quality With Stem','',true],['Red Chilli Powder — 341 Quality Stemless','',true],['Turmeric — Powder','',true]])}const s=ensure_(b,SETTINGS_SHEET,['Setting','Value']);if(s.getLastRow()===1)s.appendRow(['Show prices to customers',true]);installOrderConfirmationTrigger_();setupCatalogNotes_(c);setupOrderLayout_();setupDeliveredSheet_();return 'Setup completed without deleting existing orders or catalogue';}
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('Arjun Spice | COD Orders').addMetaTag('viewport','width=device-width,initial-scale=1')}
function getPublicCatalog(){const b=book_(),c=b.getSheetByName(CATALOG_SHEET);if(!c)throw Error('Catalog missing. Run setupArjunSpice.');const rows=c.getLastRow()>1?c.getRange(2,1,c.getLastRow()-1,3).getValues():[];const products=[],pricesPerKg={};rows.forEach(r=>{const name=String(r[0]||'').trim(),active=String(r[2]).trim().toLowerCase();if(!name||active==='false'||active==='no'||active==='0')return;products.push(name);const price=Number(r[1]);if(r[1]!==''&&Number.isFinite(price)&&price>0)pricesPerKg[name]=price});const s=b.getSheetByName(SETTINGS_SHEET);let showPrices=false;if(s&&s.getLastRow()>1){const rows=s.getRange(2,1,s.getLastRow()-1,2).getValues();const r=rows.find(x=>['showprices','show prices to customers'].includes(String(x[0]).trim().toLowerCase().replace(/\s+/g,' ')));if(r)showPrices=['true','yes','1','on'].includes(String(r[1]).trim().toLowerCase())}return {products,pricesPerKg,showPrices};}
function safe_(v,max){const s=String(v==null?'':v).trim().slice(0,max);return /^[=+@\-\t\r]/.test(s)?"'"+s:s}
function mail_(to,subject,body){if(!to)return 'No recipient configured';try{MailApp.sendEmail({to,subject,body,name:'Arjun Spice'});return 'Sent'}catch(e){console.error('Email delivery failed: '+e.message);return 'Failed: '+e.message}}
function orderRow_(sheet,id){const values=sheet.getDataRange().getValues(),idx=values[0].indexOf('Order ID');if(idx<0)throw Error('Orders sheet is missing Order ID header');for(let i=1;i<values.length;i++)if(String(values[i][idx])===id)return {row:i+1,headers:values[0],values:values[i]};return null}
function writeFields_(sheet,row,values){const h=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];Object.keys(values).forEach(k=>{const i=h.indexOf(k);if(i>=0)sheet.getRange(row,i+1).setValue(values[k])})}
function saveOrder(data){const lock=LockService.getScriptLock();lock.waitLock(20000);try{
if(!data||!Array.isArray(data.products)||!data.products.length||data.products.length>30)throw Error('Select at least one product');
const name=safe_(data.name,120),phone=String(data.phone||'').trim(),pin=String(data.pincode||'').trim();if(!name||!/^[+0-9 ()-]{10,18}$/.test(phone))throw Error('Enter valid name and phone');if(!/^\d{6}$/.test(pin))throw Error('Enter valid six-digit PIN');if(!data.address||!data.city||!data.state)throw Error('Enter full delivery address');if(data.consent!==true&&data.consent!=='on')throw Error('Consent is required');
const cat=getPublicCatalog(),allowed=new Set(cat.products);let totalKg=0,subtotal=0,missing=false;
const lines=data.products.map(p=>{const product=String(p.product||'').trim(),quantity=Number(p.quantity),unit=String(p.unit||'');if(!allowed.has(product)||!Number.isFinite(quantity)||quantity<=0||!['kg','tonnes'].includes(unit))throw Error('Invalid product or quantity');const kg=quantity*(unit==='tonnes'?1000:1);if(kg>100000000)throw Error('Quantity too large');totalKg+=kg;const price=Number(cat.pricesPerKg[product]);if(!(price>0))missing=true;return {product,kg,price:price>0?price:null,amount:price>0?Math.round(kg*price*100)/100:null,variety:safe_(p.variety,120),quantity,unit}});
if(totalKg<1)throw Error('Minimum combined order is 1 kg');if(missing)throw Error('One or more products have no price. Please ask Arjun Spice to update the Catalog sheet.');subtotal=Math.round(lines.reduce((n,p)=>n+p.amount,0)*100)/100;
const b=book_(),orders=b.getSheetByName(ORDER_SHEET);if(!orders)throw Error('Run setupArjunSpice first');columns_(orders,BASE_HEADERS.concat(EXTRA_HEADERS));const now=new Date(),id='AS-'+Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();const display=lines.map(p=>p.product+' | '+p.kg+' kg | ₹'+p.price+'/kg | ₹'+p.amount).join('\n');const owner=PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL')||'';
const dataFields={'Order ID':id,'Received At':now,'Status':'Pending COD confirmation','Full Name':name,'Business':safe_(data.business,120),'Phone':safe_(phone,18),'Email':safe_(data.email,150),'Customer Type':safe_(data.type,60),'Products (JSON)':JSON.stringify(lines),'Total Quantity (kg)':totalKg,'Full Delivery Address':[data.address,data.city,data.state,pin].map(v=>String(v||'').trim()).filter(Boolean).join(', '),'Address':safe_(data.address,500),'City':safe_(data.city,100),'State':safe_(data.state,100),'PIN':pin,'Preferred Delivery Date':safe_(data.date,20),'Packing':safe_(data.packing,100),'Payment':'Cash on Delivery','Notes':safe_(data.notes,1000),'Consent':'Yes','COD Subtotal (INR)':subtotal,'Pricing Status':'Priced at submission; delivery extra if applicable','Products (Readable)':display,'Delivery Estimate':'Approximately 5 days after confirmation; subject to delivery location','Email Notification Status':'Pending','Admin Action':'Pending'};
lines.slice(0,2).forEach((p,i)=>{const n=i+1;dataFields['Product '+n]=p.product;dataFields['Quantity '+n+' (kg)']=p.kg;dataFields['Price '+n+' (INR/kg)']=p.price;dataFields['Amount '+n+' (INR)']=p.amount});const headers=orders.getRange(1,1,1,orders.getLastColumn()).getValues()[0];orders.appendRow(headers.map(h=>dataFields[h]??''));const row=orders.getLastRow();formatOrderRow_(orders,row,lines.length);
const detail='Order ID: '+id+'\nCustomer: '+name+'\nPhone: '+phone+'\nEmail: '+(data.email||'Not provided')+'\nDelivery: '+[data.address,data.city,data.state,pin].join(', ')+'\n\nProducts:\n'+display+'\n\nTotal: ₹'+subtotal+' (shipping if applicable extra)\nPayment: COD\nStatus: Pending confirmation';const ownerStatus=mail_(owner,'New Arjun Spice COD order '+id,detail);const customerEmail=String(data.email||'').trim();const customerStatus=customerEmail?mail_(customerEmail,'Arjun Spice order received — '+id,'Order Submitted Successfully!\n\n'+detail+'\n\nEstimated delivery is around 5 days after order confirmation, subject to location and availability. This is an estimate, not a guaranteed delivery date.\nTo request cancellation, use the Cancel Order section on our website with your order ID and phone number. Cancellation is available before order processing.'): 'No customer email';writeFields_(orders,row,{'Email Notification Status':'Owner: '+ownerStatus+'; Customer: '+customerStatus});return {ok:true,orderId:id,total:subtotal,message:'Order Submitted Successfully',deliveryMessage:'Estimated delivery: around 5 days after confirmation (subject to availability and delivery location).',emailStatus:'Owner: '+ownerStatus+'; Customer: '+customerStatus};
}finally{lock.releaseLock()}}
function cancelOrder(orderId,phone){const id=String(orderId||'').trim(),mobile=String(phone||'').trim();if(!/^AS-[A-Za-z0-9-]+$/.test(id)||mobile.length<10)throw Error('Enter valid order ID and mobile number');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const b=book_(),s=b.getSheetByName(ORDER_SHEET),record=orderRow_(s,id);if(!record||String(record.values[record.headers.indexOf('Phone')]).replace(/\D/g,'')!==mobile.replace(/\D/g,''))throw Error('Order ID and mobile number do not match');const status=String(record.values[record.headers.indexOf('Status')]);if(status==='Cancelled')return {ok:true,message:'This order is already cancelled.'};if(status!=='Pending COD confirmation')throw Error('Cancellation is no longer available online. Contact Arjun Spice directly.');writeFields_(s,record.row,{'Status':'Cancelled','Cancelled At':new Date()});const owner=PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL')||'',email=String(record.values[record.headers.indexOf('Email')]||'').trim();mail_(owner,'Arjun Spice order cancelled '+id,'Order '+id+' was cancelled by the customer.');if(email)mail_(email,'Arjun Spice cancellation — '+id,'Your cancellation request was completed for order '+id+'.');return {ok:true,message:'Order Cancelled Successfully'};}finally{lock.releaseLock()}}


/** Auto-updating catalogue notes; D2 is the only formula. Existing D2:D notes are replaced. */
function setupCatalogNotes_(sheet){
  const rows=sheet.getMaxRows();
  sheet.getRange(1,4).setValue('Notes');
  if(rows>1){sheet.getRange(2,4,rows-1,1).clearContent();
    sheet.getRange(2,4).setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(B2:B="","Price not updated","Current Price: ₹"&TEXT(B2:B,"#,##0.00")&"/kg")))');}
}
/** Run once via setupArjunSpice to install the spreadsheet edit trigger. */
function installOrderConfirmationTrigger_(){
 const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='onAdminOrderEdit').forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('onAdminOrderEdit').forSpreadsheet(id).onEdit().create();
}
/** Spreadsheet menu appears after reopening the spreadsheet. */
function onOpen(){
 SpreadsheetApp.getUi().createMenu('Arjun Spice Admin')
 .addItem('Confirm selected order','confirmSelectedOrder')
 .addItem('Reject selected order','rejectSelectedOrder')
 .addItem('Mark selected order Delivered','deliverSelectedOrder')
 .addSeparator().addItem('Format Orders sheet','setupOrderLayout_')
 .addToUi();
}
function setupOrderLayout_(){
 const sheet=book_().getSheetByName(ORDER_SHEET);
 if(!sheet)return;
 const h=columns_(sheet,BASE_HEADERS.concat(EXTRA_HEADERS));
 sheet.setFrozenRows(1);
 sheet.getRange(1,1,1,h.length).setBackground('#711f19').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
 sheet.setRowHeight(1,48);
 const widths={'Order ID':180,'Received At':160,'Status':185,'Full Name':170,'Phone':150,'Email':230,'Products (Readable)':480,'Full Delivery Address':410,'COD Subtotal (INR)':155,'Total Quantity (kg)':155,'Admin Action':190,'Customer Confirmation Email Status':245};
 h.forEach((name,i)=>sheet.setColumnWidth(i+1,widths[name]||145));
 const n=Math.max(1,sheet.getMaxRows()-1);
 sheet.getRange(2,1,n,h.length).setWrap(true).setVerticalAlignment('middle');
 if(sheet.getLastRow()>1){
   sheet.setRowHeights(2,sheet.getLastRow()-1,96);
 }
 const statusCol=h.indexOf('Status')+1;
 sheet.getRange(2,statusCol,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending COD confirmation','Confirmed','Delivered','Rejected','Cancelled'],true).setAllowInvalid(false).build());
 const actionCol=h.indexOf('Admin Action')+1;
 sheet.getRange(2,actionCol,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Confirm Order','Mark Delivered','Reject Order'],true).setAllowInvalid(false).build());
 return 'Orders formatted. Use Admin Action dropdown for Confirm Order, Mark Delivered, or Reject Order.';
}
function formatOrderRow_(sheet,row,count){
 sheet.setRowHeight(row,Math.min(300,Math.max(100,50+count*32)));
 sheet.getRange(row,1,1,sheet.getLastColumn()).setWrap(true).setVerticalAlignment('middle');
}
function confirmSelectedOrder(){processSelectedOrder_('Confirmed');}
function rejectSelectedOrder(){processSelectedOrder_('Rejected');}
function deliverSelectedOrder(){processSelectedOrder_('Delivered');}
function processSelectedOrder_(state){
 const ui=SpreadsheetApp.getUi(),sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 if(sheet.getName()!==ORDER_SHEET||sheet.getActiveRange().getRow()<2){ui.alert('Select an order row in the Orders sheet first.');return;}
 const row=sheet.getActiveRange().getRow();
 const h=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
 const id=String(sheet.getRange(row,h.indexOf('Order ID')+1).getValue());
 const answer=ui.alert(state+' order '+id+'?', 'The customer will receive an email if an email address is provided.',ui.ButtonSet.YES_NO);
 if(answer!==ui.Button.YES)return;
 const result=state==='Delivered'?deliverOrder_(sheet,row):changeOrderStatus_(sheet,row,state);
 ui.alert(result);
}
function changeOrderStatus_(sheet,row,state){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const h=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const get=k=>{const i=h.indexOf(k);return i>=0?sheet.getRange(row,i+1).getValue():''};
  const id=String(get('Order ID')||'');
  if(!id)throw Error('No order ID found in selected row');
  const current=String(get('Status')||'');
  if(current==='Cancelled')throw Error('Cancelled orders cannot be confirmed or rejected');
  if(['Confirmed','Rejected'].includes(current))return 'Order '+id+' is already '+current+'. No duplicate email sent.';
  const email=String(get('Email')||'').trim();
  const detail=String(get('Products (Readable')||'');
  const fullAddress=String(get('Full Delivery Address')||[get('Address'),get('City'),get('State'),get('PIN')].filter(Boolean).join(', '));
  const message=state==='Confirmed'?
   'Your Arjun Spice Order Is Confirmed!\n\nOrder ID: '+id+'\nCustomer: '+get('Full Name')+'\n\nProducts:\n'+detail+'\n\nTotal quantity: '+get('Total Quantity (kg)')+' kg\nCOD subtotal: ₹'+get('COD Subtotal (INR)')+'\nFull delivery address: '+fullAddress+'\nPayment: Cash on Delivery\nEstimated delivery: approximately 5 days after confirmation, subject to availability and delivery location. This is not a guaranteed delivery date.\n\nThank you for choosing Arjun Spice!':
   'Your Arjun Spice order '+id+' could not be confirmed. Please contact Arjun Spice for details.';
  // Mark confirmed only after successful email, when customer has an email address.
  const result=email?mail_(email,'Arjun Spice order '+state.toLowerCase()+' — '+id,message):'Not sent: customer email missing';
  if(result!=='Sent'){
    writeFields_(sheet,row,{'Customer Confirmation Email Status':result,'Admin Action':'Pending'});
    return 'Email failed. Order remains pending: '+result;
  }
  writeFields_(sheet,row,{'Status':state,'Admin Action':state==='Confirmed'?'Confirm Order':'Reject Order','Customer Confirmation Email Status':result==='Sent'?'Sent: '+new Date().toISOString():result,...(state==='Confirmed'?{'Confirmed At':new Date()}:{})});
  return 'Order '+id+' '+state+'. Customer email: '+result;
 }finally{lock.releaseLock();}
}
function onAdminOrderEdit(e){
 if(!e||!e.range||e.range.getRow()<2||e.range.getNumRows()!==1||e.range.getNumColumns()!==1)return;
 const sheet=e.range.getSheet();if(sheet.getName()!==ORDER_SHEET)return;
 const h=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
 const col=e.range.getColumn(),v=String(e.range.getValue()).trim(),row=e.range.getRow();
 try{
  if(col===h.indexOf('Admin Action')+1){
   if(v==='Mark Delivered')deliverOrder_(sheet,row);
   else if(v==='Confirm Order'||v==='Reject Order')changeOrderStatus_(sheet,row,v==='Confirm Order'?'Confirmed':'Rejected');
  } else if(col===h.indexOf('Status')+1 && ['Confirmed','Delivered','Rejected'].includes(v)){
   // Restore the previous state before processing so that email checks cannot be bypassed.
   const previous=String(e.oldValue||'Pending COD confirmation');
   sheet.getRange(row,col).setValue(previous);
   if(v==='Delivered')deliverOrder_(sheet,row);
   else changeOrderStatus_(sheet,row,v);
  }
 }catch(err){
  console.error('Order action failed: '+err.stack);
  const statusCol=h.indexOf('Customer Confirmation Email Status')+1;
  if(statusCol>0)sheet.getRange(row,statusCol).setValue('ERROR: '+err.message);
  const actionCol=h.indexOf('Admin Action')+1;
  if(actionCol>0)sheet.getRange(row,actionCol).setValue('Pending');
 }
}
/** Copy the complete order row to Delivered Orders only after sending the delivery email. */
function deliverOrder_(sheet,row){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const h=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const get=k=>{const i=h.indexOf(k);return i<0?'':sheet.getRange(row,i+1).getValue()};
  const id=String(get('Order ID')||'');if(!id)throw Error('Missing order ID');
  if(String(get('Status'))!=='Confirmed')throw Error('Only confirmed orders can be marked delivered');
  const email=String(get('Email')||'').trim();
  if(!email){writeFields_(sheet,row,{'Customer Confirmation Email Status':'Delivery email not sent: customer email missing','Admin Action':'Pending'});return 'Customer email missing. Order not archived.';}
  const body='Your Arjun Spice order has been delivered!\n\nOrder ID: '+id+'\nCustomer: '+get('Full Name')+'\n\nProducts:\n'+get('Products (Readable')+'\n\nTotal quantity: '+get('Total Quantity (kg)')+' kg\nCOD subtotal: ₹'+get('COD Subtotal (INR)')+'\nDelivery address: '+get('Full Delivery Address')+'\nPayment method: Cash on Delivery\n\nThank you for choosing Arjun Spice!';
  const result=mail_(email,'Arjun Spice order delivered — '+id,body);
  if(result!=='Sent'){writeFields_(sheet,row,{'Customer Confirmation Email Status':'Delivery email '+result,'Admin Action':'Pending'});return 'Delivery email failed; order remains in Orders: '+result;}
  const archive=setupDeliveredSheet_();
  const ah=archive.getRange(1,1,1,archive.getLastColumn()).getValues()[0];
  const idCol=ah.indexOf('Order ID')+1;
  if(archive.getLastRow()>1 && archive.getRange(2,idCol,archive.getLastRow()-1,1).getValues().some(r=>String(r[0])===id))throw Error('Order already exists in Delivered Orders; check before retrying');
  const rowValues=sheet.getRange(row,1,1,h.length).getValues()[0];
  rowValues[h.indexOf('Status')]='Delivered';
  rowValues[h.indexOf('Admin Action')]='Mark Delivered';
  rowValues[h.indexOf('Customer Confirmation Email Status')]='Delivered email sent: '+new Date().toISOString();
  const target=archive.getLastRow()+1;
  archive.getRange(target,1,1,ah.length).setValues([ah.map(k=>{const i=h.indexOf(k);return i<0?'':rowValues[i]})]);
  archive.setRowHeight(target,Math.min(320,Math.max(100,50+String(get('Products (Readable')).split('\n').length*32)));
  archive.getRange(target,1,1,ah.length).setWrap(true).setVerticalAlignment('middle');
  SpreadsheetApp.flush();
  sheet.deleteRow(row);
  return 'Order '+id+' delivered. Email sent and order moved to Delivered Orders.';
 }finally{lock.releaseLock()}
}
function setupDeliveredSheet_(){
 const b=book_(),orders=b.getSheetByName(ORDER_SHEET);
 const h=orders.getRange(1,1,1,orders.getLastColumn()).getValues()[0].map(String);
 const s=columns_(ensure_(b,DELIVERED_SHEET,h),h);
 const archive=b.getSheetByName(DELIVERED_SHEET);
 archive.setFrozenRows(1);
 archive.getRange(1,1,1,s.length).setBackground('#375c43').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
 const widths={'Products (Readable)':480,'Full Delivery Address':410,'Email':230,'Customer Confirmation Email Status':260,'Status':165};
 s.forEach((k,i)=>archive.setColumnWidth(i+1,widths[k]||150));
 return archive;
}
/** Destructive reset: creates timestamped backup sheets before clearing historic orders.
 * Run ONLY when ready to start fresh. Does not change Catalog or Settings. */
function archiveAndClearOldOrders(){
 const b=book_(),stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'MMdd-HHmmss');
 [ORDER_SHEET,ITEM_SHEET,DELIVERED_SHEET].forEach(name=>{
  const s=b.getSheetByName(name);if(!s)return;
  if(s.getLastRow()>1){
   const backup=s.copyTo(b).setName(('Backup '+name+' '+stamp).slice(0,99));
   backup.hideSheet();
   s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).clearContent();
  }
 });
 setupOrderLayout_();
 return 'Old order data cleared from live sheets. Hidden backup sheets created with timestamp '+stamp;
}
function diagnoseCatalog(){const data=getPublicCatalog();console.log(JSON.stringify(data));return data;}
