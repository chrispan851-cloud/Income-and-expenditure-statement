(function(g){'use strict';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function docxPages(xml){
 const doc=new DOMParser().parseFromString(xml,'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('Word 文字結構無法讀取');
 const body=doc.getElementsByTagNameNS(W,'body')[0];if(!body)throw Error('Word 沒有本文');let pages=[],parts=[];
 const flush=()=>{const text=parts.join('\n').trim();if(text)pages.push(text);parts=[];};
 for(const el of body.children){if(el.localName==='sectPr')continue;
  const before=el.getElementsByTagNameNS(W,'pageBreakBefore').length;if(before)flush();
  const ps=el.localName==='p'?[el]:[...el.getElementsByTagNameNS(W,'p')];
  for(const p of ps){let line='';const walk=node=>{for(const child of node.childNodes){if(child.nodeType!==1)continue;if(child.namespaceURI===W&&child.localName==='t')line+=child.textContent;else if(child.namespaceURI===W&&child.localName==='br'){if(child.getAttributeNS(W,'type')==='page'){if(line.trim())parts.push(line);line='';flush();}else line+='\n';}else if(child.namespaceURI===W&&child.localName==='tab')line+=' ';else walk(child);}};walk(p);if(line.trim())parts.push(line);}
  const section=el.getElementsByTagNameNS(W,'sectPr')[0];if(section){const type=section.getElementsByTagNameNS(W,'type')[0]?.getAttributeNS(W,'val');flush();}
 }
 flush();return pages;
}
function inferredItem(raw){const lines=raw.normalize('NFKC').split(/\r?\n/).map(s=>s.trim());const found=[];
 for(const line of lines){const m=line.match(/^([\u3400-\u9fff][\u3400-\u9fffA-Za-z（）()、\/ \-]{1,35}?)\s*(\d[\d,]*(?:\.\d{1,2})?)\s*[Tt][Xx]?$/);if(m&&!/總計|合計|稅|小計|現金|找零|點數|序號/.test(m[1]))found.push(m[1].trim());}
 return [...new Set(found)].join('、');
}
function records(pages,rules,P){return pages.map((raw,index)=>{
 const parsed=P.parse(raw),values={direction:'expense',person:'',date:'',document:'invoice',invoice:'',expense:'',tax:'',total:'',payment:'',paymentDate:'',currency:'NTD',item:'',advance:'',note:''};
 for(const k of ['date','document','invoice','expense','tax','total'])if(parsed[k]!==undefined&&parsed[k]!==null)values[k]=String(parsed[k]);
 const matched=P.applyRules(raw,rules,values);Object.assign(values,matched.values);const inferred=!values.item?inferredItem(raw):'';if(inferred)values.item=inferred;
 const heads=raw.normalize('NFKC').split(/\r?\n/).filter(l=>/^\s*(?:電子發票證明聯|購買[票栗]品[證証]明單|免用統一發票收[據据])/.test(l)).length;
 const warnings=[...parsed.warnings,...matched.warnings];if(inferred)warnings.push('項目由商品明細預填，請確認分類與名稱。');
 if(heads>1){warnings.push('此頁可能有多張憑證，請先在掃描全能王分成一頁一張再匯出。');for(const k of ['date','invoice','expense','tax','total'])values[k]='';}
 return {raw,index,values,warnings,reviewed:false,blocked:heads>1};});}
function pdfLines(items){let rows=[];for(const item of items){if(!item.str)continue;const y=item.transform?.[5]||0,x=item.transform?.[4]||0;let row=rows.find(r=>Math.abs(r.y-y)<2);if(!row){row={y,parts:[]};rows.push(row);}row.parts.push({x,text:item.str});}return rows.sort((a,b)=>b.y-a.y).map(r=>r.parts.sort((a,b)=>a.x-b.x).map(p=>p.text).join(' ')).join('\n');}
async function extract(file){const data=await file.arrayBuffer();if(data.byteLength>25*1024*1024)throw Error('檔案超過 25 MB，請拆成較小批次');const hash=await crypto.subtle.digest('SHA-256',data),id=[...new Uint8Array(hash)].map(n=>n.toString(16).padStart(2,'0')).join('');let pages;
 if(/\.docx$/i.test(file.name)){if(!g.JSZip)throw Error('Word 讀取程式未載入，請連網重整');const zip=await g.JSZip.loadAsync(data);const entry=zip.file('word/document.xml');if(!entry)throw Error('請使用 .docx 格式');const xml=await entry.async('string');if(xml.length>15*1024*1024)throw Error('Word 文字太大，請拆分');pages=docxPages(xml);}

 else throw Error('請選掃描全能王 OCR 匯出的 Word .docx');
 if(!pages.length)throw Error('沒有可讀取的 OCR 文字，請確認匯出格式');if(pages.length>100)throw Error('超過 100 張，請分批匯入');return {pages,id};
}
const api={docxPages,inferredItem,records,pdfLines,extract};g.ReceiptBatch=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window!=='undefined'?window:globalThis);
