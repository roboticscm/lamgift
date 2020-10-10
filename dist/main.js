!function(e){var t={};function r(o){if(t[o])return t[o].exports;var s=t[o]={i:o,l:!1,exports:{}};return e[o].call(s.exports,s,s.exports,r),s.l=!0,s.exports}r.m=e,r.c=t,r.d=function(e,t,o){r.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:o})},r.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},r.t=function(e,t){if(1&t&&(e=r(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var o=Object.create(null);if(r.r(o),Object.defineProperty(o,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var s in e)r.d(o,s,function(t){return e[t]}.bind(null,s));return o},r.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return r.d(t,"a",t),t},r.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},r.p="",r(r.s=4)}([function(e,t){e.exports=require("rxjs/operators")},function(e,t){e.exports=require("xlsx")},function(e,t){e.exports=require("rxjs")},function(e,t){e.exports=require("fs")},function(e,t,r){"use strict";r.r(t);const o="dist/document/source",s="Tên thuốc, VTYT, hoá chất",n=e=>e=(e=(e=(e=(e=(e=(e=(e=(e=(e=e.toLowerCase()).replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g,"a")).replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g,"e")).replace(/ì|í|ị|ỉ|ĩ/g,"i")).replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g,"o")).replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g,"u")).replace(/ỳ|ý|ỵ|ỷ|ỹ/g,"y")).replace(/đ/g,"d")).replace(/\u0300|\u0301|\u0303|\u0309|\u0323/g,"")).replace(/\u02C6|\u0306|\u031B/g,""),c=e=>{const t=r(1),o=t.readFile(e),s=o.Sheets[o.SheetNames[0]];if(s)return t.utils.sheet_to_json(s)};var i=r(2),a=r(0);const u=(e,t)=>{const o=r(1),s=o.utils.book_new();s.Props={Title:"LamGift",Author:"Ly Van Khai 0986 409 026",Subject:"Excel Generator"},s.SheetNames.push("Summary"),s.SheetNames.push("Details");const c=r(3),u=n(t.pharmaceuticalRepresentatives.fullName);c.mkdirSync(m(),{recursive:!0});const g=[],N=[],v=[],{month:b,year:S}=h();g.push([`Tháng: ${b}/${S}`]),g.push([`Trình dược viên: ${t.pharmaceuticalRepresentatives.fullName} - ${t.pharmaceuticalRepresentatives.phoneNumber} - ${t.pharmaceuticalRepresentatives.email}`]),g.push([]),N.push([`Tháng: ${b}/${S}`]),N.push([`Trình dược viên: ${t.pharmaceuticalRepresentatives.fullName} - ${t.pharmaceuticalRepresentatives.phoneNumber} - ${t.pharmaceuticalRepresentatives.email}`]),N.push([]);const _=[],w=[];for(let r of t.data)for(let t of r.data){const o=l(e,d(r.doctor),t.TEN,t.MA);o&&(_.push({name:t.TEN,qty:o,price:t.GIA,discount:t.CK}),w.push({doctor:d(r.doctor),name:t.TEN,qty:o,price:t.GIA,discount:t.CK}))}w.sort((e,t)=>e.name>t.name?1:e.name<t.name?-1:0);const T=Object(i.from)(w).pipe(Object(a.groupBy)(e=>e.name),Object(a.mergeMap)(e=>Object(i.zip)(Object(i.of)(e.key),e.pipe(Object(a.reduce)((e,t)=>e+t.qty,0)),e.pipe(Object(a.toArray)()))));N.push(["TÊN THUỐC/BS",void 0,void 0,"SL"]),v.push({s:{r:N.length-1,c:0},e:{r:N.length-1,c:2}});let j=0;T.subscribe(e=>{N.push([e[0],void 0,void 0,e[1]]),v.push({s:{r:N.length-1,c:0},e:{r:N.length-1,c:2}});const t=e[2];t.sort(f);for(let e of t){j+=e.qty;const{firstName:t,lastName:r}=y(e.doctor);N.push([void 0,r,t,e.qty])}}),N.push(["TỔNG CỘNG",void 0,void 0,j]),v.push({s:{r:N.length-1,c:0},e:{r:N.length-1,c:2}}),_.sort((e,t)=>e.name>t.name?1:e.name<t.name?-1:0);const x=Object(i.from)(_).pipe(Object(a.groupBy)(e=>e.name),Object(a.mergeMap)(e=>e.pipe(Object(a.reduce)((e,t)=>(e.qty+=t.qty,e.price=t.price,e.discount=t.discount,e),{name:e.key,qty:0,price:void 0,discount:void 0}))),Object(a.toArray)());g.push(["TÊN","SL","ĐƠN GIÁ","TT","CK","TCK"]);let L=0,O=0,$=0;x.subscribe(e=>{for(let t of e){const e=t.qty*t.price,r=e*t.discount/100;L+=t.qty,O+=e,$+=r,g.push([t.name,t.qty,t.price,e,t.discount,r])}}),g.push(["TỔNG",L,void 0,O,void 0,$]),g.push([]),g.push([]),g.push(["File này được sinh ra bởi LamGift"]),g.push(["Powered by Lý Văn Khải - 0986 409 026 - roboticscm2018@gmail.com"]);const q=o.utils.aoa_to_sheet(g);N.push([]),N.push(["File này được sinh ra bởi LamGift"]),N.push(["Powered by Lý Văn Khải - 0986 409 026 - roboticscm2018@gmail.com"]);const C=o.utils.aoa_to_sheet(N);p(o,q,"B"),p(o,q,"C"),p(o,q,"D"),p(o,q,"E"),p(o,q,"F"),p(o,C,"D"),q["!merges"]=[],C["!merges"]=v;q["!cols"]=[{wch:30},{wch:7},{wch:7},{wch:15},{wch:5},{wch:15}];C["!cols"]=[{wch:5},{wch:20},{wch:6},{wch:10}],s.Sheets.Summary=q,s.Sheets.Details=C,o.writeFile(s,`${m()}/${u}.xlsx`)},l=(e,t,r,o)=>{const n=e.filter(e=>e["Họ - Tên Bsĩ"]===t&&(e["Mã hàng"]&&e["Mã hàng"].trim().length>0?e["Mã hàng"].trim().toLowerCase()===o.toLowerCase():e[s].toLowerCase().includes(r.toLowerCase())));return n&&n.length>0?n.map(e=>e.SL).reduce((e,t)=>e+t):void 0},h=()=>{const e=new Date,t=e.getFullYear();return{month:e.getMonth()+1,year:t}},m=()=>{const{month:e,year:t}=h();return`dist/document/dest/${t}/${e}`},p=(e,t,r)=>{const o=e.utils.decode_col(r),s=e.utils.decode_range(t["!ref"]);for(let r=s.s.r+1;r<=s.e.r;++r){const s=e.utils.encode_cell({r:r,c:o});t[s]&&("n"==t[s].t&&(t[s].z="#,##0"))}},d=e=>{const t=e.split("-");return t.length>1?t[1].trim():e},f=(e,t,r="doctor")=>{const{firstName:o,lastName:s}=g(e[r]),{firstName:n,lastName:c}=g(t[r]);return o>n?1:o<n?-1:s>c?1:s<c?-1:0},g=e=>{if(!e)return{lastName:void 0,firstName:void 0};const t=n(e),r=t.split(" ");if(r.length<2)return{lastName:void 0,firstName:t};const o=r[r.length-1].trim();r.splice(r.length-1,1);return{lastName:r.join(" ").trim(),firstName:o}},y=e=>{if(!e)return{lastName:void 0,firstName:void 0};const t=e.split(" ");if(t.length<2)return{lastName:void 0,firstName:e};const r=t[t.length-1].trim();t.splice(t.length-1,1);return{lastName:t.join(" ").trim(),firstName:r}},N=e=>e.replace(/\(.*?\)/,"").trim(),v=e=>{const t=r(1),o=t.readFile(e);if(0===o.SheetNames.length)return;const s=[];let n;for(let e in o.Sheets)if("Info"!==e){const r=t.utils.sheet_to_json(o.Sheets[e]);s.push({doctor:e,data:r})}else{const r=t.utils.sheet_to_json(o.Sheets[e]);r&&r.length>0&&(n={fullName:r[0].FullName,phoneNumber:r[0].PhoneNumber,email:r[0].Email})}return{pharmaceuticalRepresentatives:n,data:s}};new Promise((e,t)=>{const n=r(3);n.readdir(o,(i,a)=>{const l=a.filter(e=>".DS_Store"!==e&&!e.startsWith("~")),h=l.length;0===h?t({message:"Ban chua copy file goc vao thu muc: "+o,result:!1}):h>1?t({message:"Co qua nhieu file trong thu muc: "+o,result:!1}):n.readdir("dist/document/doctor_product",(n,i)=>{const a=i.filter(e=>".DS_Store"!==e&&!e.startsWith("~"));if(0===a.length)t({message:"Ban chua cau hinh trinh duoc vien trong thu muc: dist/document/doctor_product",result:!1});else{const t=c(`${o}/${l[0]}`);(e=>{const t=r(1),o=r(3);let n;o.mkdirSync(m()+"/products",{recursive:!0}),o.readdir("dist/document/price_list",(r,o)=>{const i=o.filter(e=>".DS_Store"!==e&&!e.startsWith("~")),a=i.length;if(0!==a)if(a>1)console.log("Co qua nhieu file danh muc hang hoa trong thu muc document -> price_list");else{n=c("dist/document/price_list/"+i[0]);for(let r of n){const o=e.filter(e=>!(!e["Mã hàng"]||!r["Mã thuốc"]||e["Mã hàng"].trim().toLowerCase()!==r["Mã thuốc"].trim().toLowerCase())||!!(e[s]&&r["Tên thuốc"]&&e[s].includes(r["Tên thuốc"])));if(o&&o.length>0){o.sort((e,t)=>f(e,t,"Họ - Tên Bsĩ"));const e=t.utils.book_new();e.Props={Title:"LamGift",Author:"Ly Van Khai 0986 409 026",Subject:"Excel Generator"},e.SheetNames.push("Index");const n=[];n.push(["Ngày","Họ lót BS","Tên BS","Thuốc","SL"]);let c=0;for(let e of o){const{lastName:t,firstName:r}=y(e["Họ - Tên Bsĩ"]);c+=e.SL,n.push([e["Ngày"],t,r,e[s],e.SL])}n.push(["Tổng cộng",void 0,void 0,void 0,c]);const i=t.utils.aoa_to_sheet(n),a=[{wch:20},{wch:20},{wch:10},{wch:35},{wch:10}];i["!cols"]=a;const u=[];u.push({s:{r:n.length-1,c:0},e:{r:n.length-1,c:3}}),p(t,i,"E"),i["!merges"]=u,e.Sheets.Index=i,t.writeFile(e,`${m()}/products/${N(r["Tên thuốc"])}.xlsx`)}}}else console.log("Khong tim thay file danh muc hang hoa trong thu muc document -> price_list")})})(t);for(let e of a){const r=v("dist/document/doctor_product/"+e);u(t,r)}e({result:!0,message:"Success"})}})})}).then(e=>{e.result}).catch(e=>{console.log(e.message)})}]);