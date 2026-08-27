import XLSX from "xlsx-js-style";
import type {ReviewResult} from "./analysis";

type CellStyle={font?:Record<string,unknown>;fill?:Record<string,unknown>;alignment?:Record<string,unknown>;border?:Record<string,unknown>;numFmt?:string};

const pct=(v:number|null|undefined,digits=1)=>v==null?"—":`${(v*100).toFixed(digits)}%`;
const pp=(v:number|null|undefined,base:number|null|undefined)=>v==null||base==null?"—":`${(v-base)*100>=0?"+":""}${((v-base)*100).toFixed(1)}pp`;
const process=(v:number|null|undefined,base:number|null|undefined)=>`${pct(v)} / ${pp(v,base)}`;
const team=(name:string)=>name.endsWith("团队")?name:`${name}团队`;
const address=(r:number,c:number)=>XLSX.utils.encode_cell({r,c});
const setStyle=(ws:XLSX.WorkSheet,r1:number,c1:number,r2:number,c2:number,style:CellStyle)=>{
 for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++){const cell=ws[address(r,c)];if(cell)cell.s={...(cell.s||{}),...style}}
};
const setNumFmt=(ws:XLSX.WorkSheet,r1:number,c:number,r2:number,fmt:string)=>{
 for(let r=r1;r<=r2;r++){const cell=ws[address(r,c)];if(cell)cell.z=fmt}
};

export function exportReviewWorkbook(r:ReviewResult,newcomerDays:number){
 const rows:(string|number|null)[][]=[],merges:XLSX.Range[]=[],rowHeights:{hpt:number}[]=[];
 const fullWidth=14;
 const mergeRow=(row:number,from=0,to=fullWidth)=>merges.push({s:{r:row,c:from},e:{r:row,c:to}});
 const addBlank=()=>{rows.push([]);rowHeights.push({hpt:9})};
 const addSection=(title:string)=>{const row=rows.length;rows.push([title]);mergeRow(row);rowHeights.push({hpt:24});return row};
 const addNarrative=(label:string,text:string)=>{const labelRow=rows.length;rows.push([label]);mergeRow(labelRow);rowHeights.push({hpt:20});const textRow=rows.length;rows.push([text]);mergeRow(textRow);rowHeights.push({hpt:58});return{labelRow,textRow}};
 const ltvDelta=r.current.ltv-r.previous.ltv;
 const focus=ltvDelta>15?"本期应优先复盘核心增量团队的需求诊断、方案输出、异议处理与持续追单动作，同时检查增长是否集中在少数销售，沉淀可复制打法。":ltvDelta<-15?"本期应优先抽查主要拖累团队的首次触达、到课前信任建立、完课后方案输出、异议处理和持续追单记录；新人训练、经理陪谈与SOP执行情况需结合录音和跟进记录进一步验证。":"整体基本持平不代表没有问题：需同时复盘回升基地及团队的可复制动作，与下滑基地及团队的漏斗断点，明确正负贡献如何互相抵消。";
 const totalNew=r.gradeLines.filter(x=>x.grade==="总计").reduce((s,x)=>s+x.newcomers,0);
 const totalPeople=r.gradeLines.filter(x=>x.grade==="总计").reduce((s,x)=>s+x.people,0);

 rows.push([`${r.channel} · ${r.currentPeriod} 渠道转化复盘`]);mergeRow(0);rowHeights.push({hpt:34});
 rows.push([`对比期：${r.previousPeriod}　｜　新人口径：入职${newcomerDays}天以内　｜　销售姓名去重　｜　生成时间：${new Date().toLocaleString("zh-CN",{hour12:false})}`]);mergeRow(1);rowHeights.push({hpt:22});
 addBlank();
 const kpiSection=addSection("核心指标");
 const kpiHeader=rows.length;rows.push(["整体LTV","LTV变化","到课率","出勤率","完课率","人头转化","新人占比","挂0率","整体单量","营收"]);rowHeights.push({hpt:22});
 const kpiValue=rows.length;rows.push([r.current.ltv,ltvDelta,r.current.arrival,r.current.attendance,r.current.completion,r.current.headConv,totalPeople?totalNew/totalPeople:0,r.current.people?r.current.zero/r.current.people:0,r.current.orders,r.current.revenue]);rowHeights.push({hpt:28});
 addBlank();
 const conclusionSection=addSection("一、本期结论");
 const overall=addNarrative("整体数据总结",r.summary);
 const user=addNarrative("用户侧",r.userText);
 const sales=addNarrative("销售侧",r.salesText);
 const focusRows=addNarrative("复盘重点",focus);
 addBlank();
 const gradeSection=addSection("二、分基地·分年级新人及过程数据");
 const gradeHeader=rows.length;rows.push(["基地","年级","单量","单量占比","新人数","销售人数","新人占比","LTV","出勤率"]);rowHeights.push({hpt:24});
 const gradeStart=rows.length;
 r.gradeLines.forEach(x=>{rows.push([x.base,x.grade,x.orders,x.share,x.newcomers,x.people,x.newcomerRate,x.ltv,x.attendance]);rowHeights.push({hpt:22})});
 const gradeEnd=rows.length-1;
 rows.push([`口径说明：新人按入职${newcomerDays}天以内计算；基地×年级及基地总计均按销售姓名去重，基地总计自动消除跨年级重复。`]);mergeRow(rows.length-1);rowHeights.push({hpt:24});
 addBlank();
 const managerSection=addSection("三、经理问题定位");
 const managerHeader=rows.length;rows.push(["基地","经理","年级","上期LTV","本期LTV","LTV变化","承接量","基地占比","到课/差大盘","出勤/差大盘","完课/差大盘","人头/差大盘","挂0","新人挂0","问题定位"]);rowHeights.push({hpt:28});
 const managerStart=rows.length;
 r.managers.forEach(m=>{rows.push([m.base,team(m.name),m.grade,m.previousLtv,m.ltv,m.ltv-m.previousLtv,m.orders,m.share,process(m.arrival,r.current.arrival),process(m.attendance,r.current.attendance),process(m.completion,r.current.completion),process(m.headConv,r.current.headConv),`${m.zero}/${m.people}`,`${m.newcomerZero}/${m.newcomers}`,m.diagnosis]);rowHeights.push({hpt:42})});
 const managerEnd=rows.length-1;
 addBlank();
 const cohortSection=addSection("四、连续承接与本期新承接表现");
 const cohortHeader=rows.length;rows.push(["基地","经理","连续2期人数","连续2期LTV","连续2期人头","连续2期挂0","连续3期人数","连续3期LTV","连续3期人头","连续3期挂0","新承接人数","新承接LTV","新承接人头","新承接挂0"]);rowHeights.push({hpt:28});
 const cohortStart=rows.length;
 r.managers.forEach(m=>rows.push([m.base,team(m.name),m.cohort2.people,m.cohort2.ltv,m.cohort2.headConv,`${m.cohort2.zero}/${m.cohort2.people}`,m.cohort3?.people??"数据不足",m.cohort3?.ltv??"数据不足",m.cohort3?.headConv??"数据不足",m.cohort3?`${m.cohort3.zero}/${m.cohort3.people}`:"数据不足",m.newTeam.people,m.newTeam.ltv,m.newTeam.headConv,`${m.newTeam.zero}/${m.newTeam.people}`]));
 const cohortEnd=rows.length-1;
 addBlank();
 const trendSection=addSection("五、基地近多期LTV趋势");
 const trendHeader=rows.length;rows.push(["期次",...r.bases]);rowHeights.push({hpt:24});
 const trendStart=rows.length;r.trend.forEach(t=>rows.push([t.period,...r.bases.map(b=>t.bases[b]??null)]));const trendEnd=rows.length-1;

 const ws=XLSX.utils.aoa_to_sheet(rows);
 ws["!merges"]=merges;ws["!rows"]=rowHeights;
 ws["!cols"]=[{wch:12},{wch:14},{wch:12},{wch:12},{wch:12},{wch:12},{wch:13},{wch:13},{wch:18},{wch:18},{wch:18},{wch:18},{wch:12},{wch:12},{wch:54}];
 ws["!freeze"]={xSplit:0,ySplit:6};
 const dark={patternType:"solid",fgColor:{rgb:"173E32"}},green={patternType:"solid",fgColor:{rgb:"D9EFD9"}},soft={patternType:"solid",fgColor:{rgb:"F2F7F1"}},lime={patternType:"solid",fgColor:{rgb:"DFF36B"}};
 const white={rgb:"FFFFFF"},ink={rgb:"173E32"},muted={rgb:"596861"},red={rgb:"BD4F43"};
 setStyle(ws,0,0,0,fullWidth,{fill:dark,font:{bold:true,color:white,sz:18},alignment:{horizontal:"left",vertical:"center"}});
 setStyle(ws,1,0,1,fullWidth,{fill:dark,font:{color:{rgb:"C4D5CF"},sz:10},alignment:{horizontal:"left",vertical:"center"}});
 [kpiSection,conclusionSection,gradeSection,managerSection,cohortSection,trendSection].forEach(row=>setStyle(ws,row,0,row,fullWidth,{fill:dark,font:{bold:true,color:{rgb:"DFF36B"},sz:12},alignment:{vertical:"center"}}));
 setStyle(ws,kpiHeader,0,kpiHeader,9,{fill:green,font:{bold:true,color:ink},alignment:{horizontal:"center",vertical:"center"}});
 setStyle(ws,kpiValue,0,kpiValue,9,{fill:soft,font:{bold:true,color:ink,sz:12},alignment:{horizontal:"center",vertical:"center"}});
 [overall,user,sales,focusRows].forEach(({labelRow,textRow})=>{setStyle(ws,labelRow,0,labelRow,fullWidth,{fill:green,font:{bold:true,color:ink},alignment:{vertical:"center"}});setStyle(ws,textRow,0,textRow,fullWidth,{fill:soft,font:{color:muted,sz:11},alignment:{wrapText:true,vertical:"top"}})});
 [gradeHeader,managerHeader,cohortHeader,trendHeader].forEach(row=>setStyle(ws,row,0,row,fullWidth,{fill:green,font:{bold:true,color:ink},alignment:{horizontal:"center",vertical:"center",wrapText:true}}));
 setStyle(ws,gradeStart,0,gradeEnd,8,{alignment:{horizontal:"center",vertical:"center"},border:{bottom:{style:"thin",color:{rgb:"D9E2DD"}}}});
 setStyle(ws,managerStart,0,managerEnd,14,{alignment:{vertical:"center",wrapText:true},border:{bottom:{style:"thin",color:{rgb:"D9E2DD"}}}});
 setStyle(ws,cohortStart,0,cohortEnd,13,{alignment:{horizontal:"center",vertical:"center"},border:{bottom:{style:"thin",color:{rgb:"D9E2DD"}}}});
 setStyle(ws,trendStart,0,trendEnd,Math.max(0,r.bases.length),{alignment:{horizontal:"center",vertical:"center"}});
 for(let row=gradeStart;row<=gradeEnd;row++)if(rows[row][1]==="总计")setStyle(ws,row,0,row,8,{fill:soft,font:{bold:true,color:ink}});
 for(let row=managerStart;row<=managerEnd;row++){const delta=Number(rows[row][5]);setStyle(ws,row,5,row,5,{font:{bold:true,color:delta<0?red:ink}})}
 setNumFmt(ws,kpiValue,0,kpiValue,"0.0");setNumFmt(ws,kpiValue,1,kpiValue,"+0.0;-0.0;0.0");[2,3,4,5,6,7].forEach(c=>setNumFmt(ws,kpiValue,c,kpiValue,"0.0%"));setNumFmt(ws,kpiValue,8,kpiValue,"#,##0");setNumFmt(ws,kpiValue,9,kpiValue,"#,##0");
 [3,4,5].forEach(c=>setNumFmt(ws,managerStart,c,managerEnd,c===5?"+0.0;-0.0;0.0":"0.0"));setNumFmt(ws,managerStart,7,managerEnd,"0.0%");
 [3,7].forEach(c=>setNumFmt(ws,gradeStart,c,gradeEnd,c===3?"0.0%":"0.0"));[6,8].forEach(c=>setNumFmt(ws,gradeStart,c,gradeEnd,"0.0%"));
 [3,4,7,8,11,12].forEach(c=>setNumFmt(ws,cohortStart,c,cohortEnd,c===4||c===8||c===12?"0.0%":"0.0"));
 if(trendEnd>=trendStart)for(let c=1;c<=r.bases.length;c++)setNumFmt(ws,trendStart,c,trendEnd,"0.0");
 ws["!autofilter"]={ref:`A${managerHeader+1}:O${managerEnd+1}`};

 const managerRows=r.managers.map((m,i)=>({优先级:i+1,基地:m.base,经理:team(m.name),年级:m.grade,连续两期拖累:m.repeatDrag?"是":"否",承接量:m.orders,基地单量占比:m.share,上期LTV:m.previousLtv,本期LTV:m.ltv,LTV变化:m.ltv-m.previousLtv,到课率:m.arrival,到课较大盘:m.arrival==null||r.current.arrival==null?null:m.arrival-r.current.arrival,出勤率:m.attendance,出勤较大盘:m.attendance==null||r.current.attendance==null?null:m.attendance-r.current.attendance,完课率:m.completion,完课较大盘:m.completion==null||r.current.completion==null?null:m.completion-r.current.completion,人头转化:m.headConv,人头较大盘:m.headConv-r.current.headConv,挂0人数:m.zero,团队人数:m.people,新人数:m.newcomers,新人挂0:m.newcomerZero,Top2营收占比:m.top2Share,连续2期人数:m.cohort2.people,连续2期LTV:m.cohort2.ltv,连续2期人头:m.cohort2.headConv,连续2期挂0:m.cohort2.zero,连续3期人数:m.cohort3?.people??null,连续3期LTV:m.cohort3?.ltv??null,连续3期人头:m.cohort3?.headConv??null,连续3期挂0:m.cohort3?.zero??null,本期新承接人数:m.newTeam.people,本期新承接LTV:m.newTeam.ltv,本期新承接人头:m.newTeam.headConv,本期新承接挂0:m.newTeam.zero,问题定位:m.diagnosis}));
 const managerWs=XLSX.utils.json_to_sheet(managerRows);managerWs["!freeze"]={xSplit:2,ySplit:1};managerWs["!autofilter"]={ref:managerWs["!ref"]||"A1:AJ1"};managerWs["!cols"]=Object.keys(managerRows[0]||{经理:""}).map(k=>({wch:k==="问题定位"?56:Math.max(11,Math.min(18,k.length*2+2))}));
 if(managerRows.length){setStyle(managerWs,0,0,0,Object.keys(managerRows[0]).length-1,{fill:dark,font:{bold:true,color:white},alignment:{horizontal:"center",vertical:"center",wrapText:true}})}
 const managerLast=managerRows.length;[6,10,11,12,13,14,15,16,17,22,25,29,33].forEach(c=>setNumFmt(managerWs,1,c,managerLast,"0.0%"));[7,8,9,24,28,32].forEach(c=>setNumFmt(managerWs,1,c,managerLast,c===9?"+0.0;-0.0;0.0":"0.0"));

 const trendRows=r.trend.map(t=>({期次:t.period,...Object.fromEntries(r.bases.map(b=>[b,t.bases[b]??null]))}));
 const trendWs=XLSX.utils.json_to_sheet(trendRows);trendWs["!freeze"]={xSplit:1,ySplit:1};trendWs["!autofilter"]={ref:trendWs["!ref"]||"A1:B1"};trendWs["!cols"]=[{wch:14},...r.bases.map(()=>({wch:14}))];
 setStyle(trendWs,0,0,0,r.bases.length,{fill:dark,font:{bold:true,color:white},alignment:{horizontal:"center",vertical:"center"}});
 if(trendRows.length)for(let c=1;c<=r.bases.length;c++)setNumFmt(trendWs,1,c,trendRows.length,"0.0");

 const wb=XLSX.utils.book_new();wb.Props={Title:`${r.channel} ${r.currentPeriod} 渠道转化复盘`,Subject:"教育销售渠道转化复盘",Author:"胜利",CreatedDate:new Date()};wb.Workbook={CalcPr:{calcMode:"auto"}};
 XLSX.utils.book_append_sheet(wb,ws,"渠道转化复盘");
 XLSX.utils.book_append_sheet(wb,managerWs,"经理完整明细");
 XLSX.utils.book_append_sheet(wb,trendWs,"趋势数据");
 XLSX.writeFile(wb,`${r.channel}_${r.currentPeriod}_渠道转化复盘.xlsx`,{cellStyles:true});
}
