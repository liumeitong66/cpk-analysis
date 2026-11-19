async function buildAnalysisFiles(canvas,rows){
const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const main='http://schemas.openxmlformats.org/spreadsheetml/2006/main', rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships', pkg='http://schemas.openxmlformats.org/package/2006/relationships';
const png=await new Promise(r=>canvas.toBlob(r,'image/png'));if(!png)throw Error('图表生成失败');
const sheet=rows.map((row,i)=>`<row r="${i+1}">${row.map((v,j)=>`<c r="${String.fromCharCode(65+j)}${i+1}" ${typeof v==='number'&&isFinite(v)?`><v>${v}</v>`:`t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is>`}</c>`).join('')}</row>`).join('');
const files={
'[Content_Types].xml':`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${[1,2].map(n=>`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`,
'_rels/.rels':`<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
'xl/workbook.xml':`<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="分析数据" sheetId="1" r:id="rId1"/><sheet name="完整报告图表" sheetId="2" r:id="rId2"/></sheets></workbook>`,
'xl/_rels/workbook.xml.rels':`<Relationships xmlns="${pkg}">${[1,2].map(n=>`<Relationship Id="rId${n}" Type="${rel}/worksheet" Target="worksheets/sheet${n}.xml"/>`).join('')}</Relationships>`,
'xl/worksheets/sheet1.xml':`<worksheet xmlns="${main}"><cols><col min="1" max="1" width="38" customWidth="1"/><col min="2" max="2" width="55" customWidth="1"/><col min="3" max="22" width="24" customWidth="1"/></cols><sheetData>${sheet}</sheetData></worksheet>`,
'xl/worksheets/sheet2.xml':`<worksheet xmlns="${main}" xmlns:r="${rel}"><sheetData/><drawing r:id="rId1"/></worksheet>`,
'xl/worksheets/_rels/sheet2.xml.rels':`<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
'xl/drawings/drawing1.xml':`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="10477500" cy="${Math.round(canvas.height/canvas.width*10477500)}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="完整分析报告"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${rel}" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`,
'xl/drawings/_rels/drawing1.xml.rels':`<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/image" Target="../media/report.png"/></Relationships>`,
'xl/media/report.png':new Uint8Array(await png.arrayBuffer())};
return files;
}
async function downloadAnalysisWorkbook(items,filename){
const files={};let content='';let sheets='';let links='';
for(let i=0;i<items.length;i++){
 const n=i+1,item=items[i];
 const chartHeight=Math.round(item.canvas.height/item.canvas.width*1100);
 const blankRows=Array.from({length:Math.ceil(chartHeight/20)+3},()=>[]);
 const source=await buildAnalysisFiles(item.canvas,[...blankRows,...item.rows]);
 if(!i){files['_rels/.rels']=source['_rels/.rels'];content=source['[Content_Types].xml'].replace(/<Override PartName="\/xl\/(?:worksheets|drawings)[\s\S]*?(?=<\/Types>)/,'').replace('</Types>','');}
 content+=`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
 const escape=v=>String(v).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
 sheets+=`<sheet name="${escape(item.name)}" sheetId="${n}" r:id="rId${n}"/>`;
 links+=`<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`;
 files[`xl/worksheets/sheet${n}.xml`]=source['xl/worksheets/sheet1.xml'].replace('<worksheet xmlns=', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns=').replace('</worksheet>','<drawing r:id="rId1"/></worksheet>');
 files[`xl/worksheets/_rels/sheet${n}.xml.rels`]=source['xl/worksheets/_rels/sheet2.xml.rels'].replace('drawing1.xml',`drawing${n}.xml`);
 files[`xl/drawings/drawing${n}.xml`]=source['xl/drawings/drawing1.xml'];
 files[`xl/drawings/_rels/drawing${n}.xml.rels`]=source['xl/drawings/_rels/drawing1.xml.rels'].replace('report.png',`report${n}.png`);
 files[`xl/media/report${n}.png`]=source['xl/media/report.png'];
}
files['[Content_Types].xml']=content+'</Types>';
files['xl/workbook.xml']=`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets></workbook>`;
files['xl/_rels/workbook.xml.rels']=`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${links}</Relationships>`;
const enc=new TextEncoder(),parts=[],central=[];let offset=0;
const header=n=>{const a=new Uint8Array(n);return [a,new DataView(a.buffer)];};
for(const [path,value] of Object.entries(files)){
const name=enc.encode(path),data=typeof value==='string'?enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+value):value;
let crc=0xffffffff;for(const byte of data){crc^=byte;for(let k=0;k<8;k++)crc=(crc>>>1)^((crc&1)?0xedb88320:0);}crc=(crc^0xffffffff)>>>0;
const [local,l]=header(30);l.setUint32(0,0x04034b50,true);l.setUint16(4,20,true);l.setUint32(14,crc,true);l.setUint32(18,data.length,true);l.setUint32(22,data.length,true);l.setUint16(26,name.length,true);parts.push(local,name,data);
const [c,d]=header(46);d.setUint32(0,0x02014b50,true);d.setUint16(4,20,true);d.setUint16(6,20,true);d.setUint32(16,crc,true);d.setUint32(20,data.length,true);
d.setUint32(24,data.length,true);d.setUint16(28,name.length,true);d.setUint32(42,offset,true);central.push(c,name);offset+=30+name.length+data.length;
}
const [end,e]=header(22);e.setUint32(0,0x06054b50,true);e.setUint16(8,Object.keys(files).length,true);e.setUint16(10,Object.keys(files).length,true);e.setUint32(12,central.reduce((s,a)=>s+a.length,0),true);e.setUint32(16,offset,true);
const url=URL.createObjectURL(new Blob([...parts,...central,end],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const a=document.createElement('a');a.href=url;a.download=filename+'.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}