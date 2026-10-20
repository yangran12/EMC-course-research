import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const workspaceDir='D:/PycharmProjects/电磁兼容';
const SKILL_DIR='C:/Users/Lenovo/.codex/plugins/cache/openai-primary-runtime/presentations/26.1017.11052/skills/presentations';
const RUNTIME_ROOT='C:/Users/Lenovo/.cache/codex-runtimes/codex-primary-runtime/dependencies';
const RUNTIME_PYTHON=`${RUNTIME_ROOT}/python/python.exe`;
process.env.RUNTIME_NODE_MODULES=`${RUNTIME_ROOT}/node/node_modules`;
const requireRuntime=createRequire(`${process.env.RUNTIME_NODE_MODULES}/__runtime__.cjs`);
const {Presentation,PresentationFile}=await import(pathToFileURL(requireRuntime.resolve('@oai/artifact-tool')).href);
const {finalizePresentation,resolvePresentationFont,applyPresentationChartFont}=await import(pathToFileURL(`${SKILL_DIR}/container_tools/artifact_tool_utils.mjs`).href);
const font=resolvePresentationFont({fontFamily:'Microsoft YaHei',availableFonts:['Microsoft YaHei']});
const fontPolicy={basis:'design',families:[font]};
const assetDir=`${workspaceDir}/演示源文件/磁环素材`;
const buildDir=`${workspaceDir}/.build/ferrite`;
const finalName=process.argv[2]||'南师大风格_电源线上的小鼓包_磁环抑制干扰.pptx';
if(path.basename(finalName)!==finalName||!finalName.endsWith('.pptx'))throw new Error('请使用尚不存在的PPTX输出文件名');
const finalPath=`${workspaceDir}/输出/${finalName}`;
const C={blue:'#002060',accent:'#133FAD',red:'#C00000',text:'#222B3A',muted:'#5B6473',line:'#CBD3DF',white:'#FFFFFF'};
const p=Presentation.create({slideSize:{width:1280,height:720}});
const slides=[];
const citations={
 r1:'Würth Elektronik. Cable Solutions, Product Navigator. 访问日期2026-10-20. https://www.we-online.com/en/components/productnavigator/cable-solutions',
 r2:'Murata Manufacturing. Conductor conduction and common mode. Guidelines for EMI Suppression (EMC), Section 1, Chapter 5. Fig.5-2-3与Fig.5-2-10. https://www.murata.com/products/emc/emifil/library/knowhow/basic/chapter05-p1',
 r3:'Fair-Rite Products Corp. Catalog, 17th Edition. Technical Information,印刷页144 Fig.2（Z/R/X），印刷页152 Fig.25（1/2/3匝）. https://www.fair-rite.com/wp-content/uploads/2023/05/Fair-Rite_Catalog_17th_Edition.pdf',
 r4:'Fair-Rite Products Corp. Round Cable Snap-It, Part 0431164281. 原厂典型阻抗表，单匝器件测试. 访问日期2026-10-20. https://www.fair-rite.com/printer_friendly_datasheet.php?part=0431164281',
 r5:'Murata Manufacturing. Noise Suppression Measures at DC Power Cable Connection. Guidelines for EMI Suppression (EMC), Section 2, Theme 7. 7-2-1与7-2-4，Fig.2-1、Fig.2-14最左面板. https://www.murata.com/en-global/products/emc/emifil/library/knowhow/basic/s2-chapter07-p1'
};
function text(s,v,x,y,w,h,size=23,color=C.text,bold=false,align='left'){
 const sh=s.shapes.add({geometry:'textbox',position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 sh.text=v;sh.text.style={typeface:font,fontSize:size,color,bold,alignment:align,verticalAlignment:'top',wrap:'square',autoFit:'none',insets:0};return sh;
}
function para(s,v,x,y,w,h,size=23){
 const sh=text(s,v.replaceAll('**',''),x,y,w,h,size);
 sh.text=[{runs:v.split(/(\*\*.*?\*\*)/g).filter(Boolean).map(a=>({run:a.startsWith('**')?a.slice(2,-2):a,style:{typeface:font,fontSize:`${size}px`,bold:a.startsWith('**'),color:a.startsWith('**')?C.blue:C.text}}))}];return sh;
}
function heading(s,v,x,y,w=550,size=27){return text(s,v,x,y,w,42,size,C.accent,true);}
function rule(s,x,y,w,color=C.line,width=1){s.shapes.add({geometry:'line',position:{left:x,top:y,width:w,height:0},fill:'none',line:{style:'solid',fill:color,width}});}
async function img(s,file,x,y,w,h,alt){
 const b=await fs.readFile(file.startsWith('D:')?file:`${assetDir}/${file}`);
 s.images.add({blob:b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength),contentType:'image/png',alt,fit:'contain',position:{left:x,top:y,width:w,height:h}});
}
async function slide(title,source){
 const s=p.slides.add();slides.push(s);s.background.fill=C.white;
 text(s,title,60,43,867,58,40,C.blue,true);
 await img(s,`${workspaceDir}/演示源文件/素材/南师大校标.png`,956,28,264,50.66,'南京师范大学校标，沿用用户参考答辩稿');
 rule(s,60,118,1160,C.accent,2);text(s,source,60,668,1048,32,16,C.muted);
 text(s,String(slides.length).padStart(2,'0'),1130,668,90,32,16,C.muted,false,'right');return s;
}
function table(s,values,widths,x,y,w,h,size=22){
 const t=s.tables.add({rows:values.length,columns:values[0].length,left:x,top:y,width:w,height:h,columnWidths:widths,values});
 t.borders.assign({style:'solid',fill:C.line,width:1});
 for(let r=0;r<values.length;r++)for(let c=0;c<values[0].length;c++){
  const cell=t.getCell(r,c);cell.fill=r===0?C.blue:(r%2?C.white:'#F4F7FA');
  cell.text.style={typeface:font,fontSize:size,color:r===0?C.white:C.text,bold:r===0,verticalAlignment:'middle',wrap:'square',autoFit:'none'};
  cell.text.insets={left:14,right:12,top:8,bottom:8};
 }t.rows[0].height=44;return t;
}
function notes(s,seconds,talk,refs,extra=''){
 s.speakerNotes.text=`【建议讲解 ${seconds} 秒】\n${talk}\n\n【资料来源】\n${refs}\n\n【答疑补充】\n${extra||'如需展开，先回到本页图示解释，再说明结论成立的条件。'}`;
}

// 01. Minimal school-style cover, keeping the user's smaller text preference.
{
 const s=p.slides.add();slides.push(s);s.background.fill=C.white;
 await img(s,`${workspaceDir}/演示源文件/素材/南师大校标.png`,345,70,590,113.2,'南京师范大学组合校标');
 text(s,'电源线上的“小鼓包”有什么用？',90,270,1100,90,48,C.blue,true,'center');
 text(s,'磁环抑制电磁干扰的原理与应用',130,382,1020,56,32,C.accent,true,'center');
 text(s,'电磁兼容课程文献调研',160,484,960,43,26,C.blue,false,'center');
 text(s,'汇报人：杨冉    授课老师：周孟夏',160,554,960,40,24,C.blue,false,'center');
 notes(s,20,'老师、同学们好。大家可能见过某些电源线上的圆柱形小鼓包，里面常装有铁氧体磁芯。今天我们围绕它讲一个具体问题：磁环怎样帮助减轻线缆上的电磁干扰，为什么有时有效、有时效果有限。',citations.r1);
}
// 02. Identify the actual component before introducing terminology.
{
 const s=await slide('线缆磁环：实物与用途','实物图：[4] Fair-Rite；用途依据 [1]、[2]，以二线直流供电线为例');
 heading(s,'塑料外壳内的铁氧体磁芯',65,153,635);
 await img(s,'ferrite_snap_photo.png',65,207,643,300,'原厂夹扣式线缆磁环实物');
 para(s,'夹扣式磁环可以打开，将线缆放入后闭合。线缆从磁芯孔中穿过，磁环与导线通过磁场发生作用，**不需要接入电路节点**。',65,540,643,96);
 heading(s,'为什么会装在线缆上？',760,153,455);
 para(s,'电子设备的高频噪声可能沿电源线或数据线传播，线缆也可能向外辐射。磁环是针对这类路径的一种抑制措施，选用时要结合实际频段与电流模式。',760,203,455,138);
 heading(s,'本次调研的具体问题',760,373,455);
 para(s,'统一观察两根直流供电线：正常供电电流怎样经过磁环？共模干扰电流为何受到不同影响？我们再通过器件数据与一个厂商测量案例检验这些解释。',760,423,455,148);
 text(s,'先看电流怎样走，再讨论磁环是否合适。',760,600,455,42,24,C.red,true);
 notes(s,60,'先让大家认出实物。它的外壳只是便于固定，内部的铁氧体磁芯才是发挥作用的部分。导线不必剥开，也没有与磁环形成新的电气接点。后面统一用两根直流供电线举例，避免把多种线缆接法一起讲。',citations.r1+'\n'+citations.r2+'\n'+citations.r4,'并非所有线缆都需要外加磁环。产品也可能通过板内滤波、布局或屏蔽等方式处理干扰，是否采用取决于实际EMC需求。');
}
// 03. Actual board context links ordinary power transfer with high-frequency noise.
{
 const s=await slide('干扰怎样传到电源线？','图：[5] Fig.2-1；原厂试验板用于说明供电与高频电路并存的场景');
 heading(s,'一个真实的数字电路试验板',65,153,645);
 await img(s,'murata_dc_cable_fig2_01_original.png',65,209,643,343,'村田实验板正反面与10cm×7cm尺寸，带DC-DC和数字电路');
 para(s,'原厂装置通过直流电源线输入 5 V，再提供 3.3 V 给数字电路；板上有 40 MHz 数字活动。供电线连接着设备，也可能成为噪声传播的路径。',65,570,643,82,22);
 heading(s,'电源线里不只有直流电流',760,153,455);
 para(s,'快速开关的电压、电流会激励高频噪声。它可通过实际连线或寄生电容耦合进入线缆，再通过设备和周围环境形成的路径返回。',760,203,455,131);
 heading(s,'本次关注共模分量',760,369,455);
 para(s,'线缆上的电流可以分解成不同模式。接下来的图只观察同一个截面上的电流方向，用它判断整体穿过磁环后，磁化作用是抵消还是相加。',760,420,455,135);
 text(s,'线缆既传送能量，也可能传送干扰。',760,600,455,42,24,C.red,true);
 notes(s,65,'左边是村田的真实试验板，只需要记住它用直流线供电，同时板上存在快速变化的数字信号。我们不展开芯片细节。关键是直流供电正常，不代表线上只有直流。高频噪声仍可能通过导线和寄生耦合形成回路。下一页用电流方向解释两种模式。',citations.r5+'\n'+citations.r2,'40 MHz是该试验板的数字时钟，不等于所有干扰只出现在40 MHz。实际噪声还可能含谐波，案例频谱由原厂测量给出。');
}
// 04. Manufacturer's original current-direction diagram, with accessible Chinese reading cues.
{
 const s=await slide('两种电流方向：往返与共模','图：[2] Fig.5-2-3；Normal mode 在本文称为差模形式');
 await img(s,'chapter05-p1_img0004.png',65,159,1150,321,'村田正常模式和共模电流原图，含外部寄生电容返回路径');
 heading(s,'正常供电：一根去，一根回',65,510,551);
 para(s,'从同一截面观察，去线和回线的工作电流方向相反。两根线一起穿过磁环时，磁化作用大致抵消，因此正常供电受到的影响较小。',65,556,551,91,22);
 heading(s,'共模分量：两根线同向',660,510,555);
 para(s,'共模分量在两根线上同向，磁化作用相加。电流通过外部寄生路径返回，所以整体穿芯的磁环可以对该模式增加阻抗。',660,556,555,91,22);
 notes(s,95,'请先看左图的两根线：上面流出去，下面流回来。不要只追一根线，要把去回两根作为整体。右图的共模分量在两根线上同向，但电流仍需要回路，返回路径可能经寄生电容和环境。方向不同，使两种模式对同一个磁环的磁化作用不同。',citations.r2,'差模形式不一定就是有用信号，也可能包含差模噪声。本页只是用正常直流供电说明一去一回的方向；真实总电流是工作电流和噪声等分量的叠加。');
}
// 05. Frequency-dependent loss and reactance, explained without material equations.
{
 const s=await slide('磁环怎样减小高频干扰？','图：[3] 印刷页144 Fig.2；原图为43材料示例，区别于下一页的31材料产品');
 heading(s,'磁环阻抗随频率变化',65,153,643);
 await img(s,'Fair_Rite_Fig02_ZRX.png',82,205,610,367,'Fair-Rite原图：磁环总阻抗Z、电阻分量R与电抗分量X随频率变化');
 para(s,'图中 Z 是总阻抗幅值，R 对应损耗分量，X 对应电抗分量。读图先看横轴频率，再判断目标频段内阻抗是否足够。',65,590,643,58,22);
 heading(s,'对共模电流增加阻抗',760,153,455);
 para(s,'当线缆整体穿过磁环，共模电流的磁化作用相加。在合适频段，磁环使这条噪声路径更难通过，从而有机会减轻沿线传播和由线缆引起的辐射。',760,203,455,139);
 heading(s,'部分高频能量被耗散',760,374,455);
 para(s,'铁氧体具有频率相关的损耗，其中电阻性部分可把部分噪声能量转化为热。磁环的表现同时包含损耗与电抗效应，不能用一个理想电感值概括全部频段。',760,424,455,145);
 text(s,'效果取决于频段、穿线方式和实际回路。',760,600,455,42,24,C.red,true);
 notes(s,75,'这页只解释两个词。阻抗表示电流通过这条路径时受到的作用；损耗表示部分高频能量转成热。左图的三条曲线随频率变化，说明磁环的作用不是恒定的。大家不需要推导材料公式，只要能按频率读曲线，并联系上一页的共模电流。',citations.r3+'\n'+citations.r1,'图中的Z是总阻抗幅值，不能把R和X的幅值直接算术相加。磁环不会识别什么是有用信号，模式与频率决定它对电流的影响。');
}
// 06. Editable, exact manufacturer data. This is component impedance, not a DUT attenuation result.
{
 const s=await slide('看数据选磁环：先看干扰频段','数据：[4] Fair-Rite 0431164281，原厂单匝典型阻抗；不是整机噪声测试');
 const data=JSON.parse(await fs.readFile(`${workspaceDir}/演示源文件/磁环_图表数据.json`,'utf8'));
 const ch=s.charts.add('bar',{position:{left:54,top:158,width:706,height:382},categories:data.frequency_MHz.map(f=>`${f} MHz`),series:[{name:'典型阻抗',values:data.typical_impedance_ohm,fill:C.accent}],barOptions:{direction:'column',grouping:'clustered',gapWidth:90},hasLegend:false,xAxis:{textStyle:{fontSize:19},majorGridlines:null},yAxis:{title:'典型阻抗 / Ω',min:0,max:450,majorUnit:100,numberFormatCode:'0',textStyle:{fontSize:19},majorGridlines:{fill:'#DDE3EA',width:1}},dataLabels:{showValue:true,position:'outEnd',numberFormatCode:'0',textStyle:{fontSize:21}},chartFill:'none',plotAreaFill:'none'});
 applyPresentationChartFont(ch,{fontFamily:font});
 para(s,'同一个磁环在不同频率下具有不同阻抗。柱图列出原厂表格的 6 个频点，供比较频率响应；柱间不代表线性变化，也不外推到更高频段。',65,573,668,82,22);
 heading(s,'先知道噪声在哪里',796,153,421);
 para(s,'选型应先确定主要干扰频段，再查该频段的阻抗曲线。图中的 100 MHz、315 Ω 表示器件的典型阻抗，实际噪声改善量还受源、负载和返回路径影响。',796,203,421,147);
 heading(s,'数据手册要读清测试条件',796,387,421);
 para(s,'本例是单匝器件阻抗测试，原厂还区分典型值与受控测试频点。更换材料、尺寸或穿线次数后，需要重新看相应曲线，不能只比较外形大小。',796,437,421,144);
 text(s,'器件曲线用于选型，效果靠实测验证。',796,598,421,60,23,C.red,true);
 notes(s,60,'这里选一个型号只是为了让数据有明确来源，不要求记住型号。请看同一器件从1到250兆赫的典型阻抗变化。315欧是阻抗，不是315分贝。器件数据帮助我们选候选磁环，最终能降低多少噪声还要做前后对比。',citations.r4,'六个频点：1/5/10/25/100/250 MHz，对应20/75/104/170/315/380 Ω。均为原厂典型值，不是本组实测。原厂在标记+的频点给出最小值控制说明，本汇报未把典型值称作保证值。');
}
// 07. Actual turns curve makes a simple, evidence-based engineering comparison.
{
 const s=await slide('多穿几次，效果一定更好吗？','图：[3] 印刷页152 Fig.25；43材料、14/6/28 mm原厂示例，区别于第6页产品');
 heading(s,'同一磁环的 1、2、3 匝曲线',65,153,643);
 await img(s,'Fair_Rite_Fig25_Turns_1_2_3.png',91,207,596,366,'Fair-Rite同一磁环1/2/3匝阻抗-频率原图');
 para(s,'直穿磁芯孔一次已经算 1 匝。增加穿孔次数后，阻抗曲线和峰值频率都会改变，图中不同曲线在高频区还会交叉。',65,590,643,58,22);
 heading(s,'增加匝数有作用条件',760,153,455);
 para(s,'在图示的部分频段，多匝可以提高阻抗；但效果不能推广到所有频率。绕线引入的寄生参数会改变频率响应，因此“多绕几圈”需要结合目标频段判断。',760,203,455,140);
 heading(s,'去线、回线作为一束穿过',760,382,455);
 para(s,'当目标是共模抑制时，两根供电线应一起穿同一个磁环，穿过次数相同。这样正常供电电流的磁化作用仍可大致抵消，共模分量的作用则相加。',760,432,455,139);
 text(s,'按目标频段选匝数，再比较实际效果。',760,600,455,42,24,C.red,true);
 notes(s,65,'先指出一条线直穿孔一次已经是一匝。再看图：低一些的频率中，多匝的阻抗更高，但到了高频，曲线会交叉。这就说明匝数并非越多越好。穿线时去线和回线应作为一束一起处理，不能只为了提高匝数而把工作电流的返回线丢在外面。',citations.r3+'\n'+citations.r2,'图为原厂43材料、外径/内径/长度14/6/28 mm示例。不能与第6页31材料型号合并成同一器件的实验。匝数效应受频率与寄生参数影响，不使用无条件N²估算。');
}
// 08. Actual installation pictures and a short native checklist rather than advanced derivation.
{
 const s=await slide('磁环的位置与接法','图：[2] Fig.5-2-10，左为线缆绕环，右为柔性板夹磁芯；方法依据 [1]–[3]');
 await img(s,'chapter05-p1_img0011.png',65,153,1150,242,'村田线缆绕磁环及夹式磁芯的实际使用示例');
 table(s,[['检查项目','为什么要检查'],['去线与回线一起穿芯','保持正常往返电流的磁化作用大致抵消，针对共模分量增加阻抗'],['从线缆接口附近开始比较位置','减少高频噪声进入外接线缆的机会，具体位置结合源和返回路径调整'],['闭合到位，固定并保持布线一致','夹扣闭合、线缆走向及附近导体会影响结果，前后测量需保持可比']],[350,800],65,425,1150,171,21.5);
 para(s,'从噪声进入或离开线缆的接口附近开始试，结合实际回路调整位置，并保持前后测试布置一致。',65,618,1150,40,22);
 notes(s,60,'照片是原厂的两种使用方式，不是本组搭建的实验。接法看三点：去回线一起穿；从噪声进入或离开线缆的接口附近开始试；闭合和布线保持一致。接口附近是一个起点，最终还要看真实路径与测试效果。',citations.r2+'\n'+citations.r1+'\n'+citations.r3,'不同设备可能有多个噪声源和不同返回路径，所以没有通用的最佳安装距离。图中的具体匝数不能替代目标频段曲线与实测。');
}
// 09. One original before/after case; no invented trace digitization or test configuration.
{
 const s=await slide('原厂案例：加磁环前后的辐射噪声','图：[5] Fig.2-14 最左面板；粉色虚线为加装前，蓝色曲线为加装后');
 heading(s,'同一装置的加装前后对比',65,153,680);
 await img(s,'murata_fig2_14_left_spectrum.png',65,211,681,396,'村田原厂左面板辐射频谱，保留Before/After、MHz及dBμV/m单位');
 text(s,'横轴：频率 MHz    纵轴：辐射电平 dBμV/m',65,620,681,34,20,C.muted);
 heading(s,'原厂报告的结果',795,153,421);
 text(s,'部分峰值处降低超过 10 dB',795,204,421,79,30,C.red,true);
 para(s,'在线缆上加装磁芯后，部分峰值处辐射电平下降。原文仍指出存在较强辐射，说明线缆之外还需要考虑板级噪声。',795,303,421,137);
 heading(s,'比较依赖相同测试条件',795,466,421);
 para(s,'案例装置：5 V 输入、3.3 V 数字供电、40 MHz 数字电路。暗室内测距 3 m，天线高 1 m，水平极化；结果限于该装置与测量工况。',795,516,421,121,22);
 notes(s,85,'这是真实原厂案例，不是我们测得的数据。粉色虚线是加装前，蓝线是加装后，纵轴测的是辐射电平。原厂文字说峰值处降低超过10分贝，我们只采用这个有来源的表述，不从小图新增精确峰值。曲线没有全部消失，作者也指出电路板仍可能辐射，所以线缆磁环只是抑制路径之一。',citations.r5,'只取Fig.2-14最左面板，未混用后两面板附加50cm GND或V+导线的工况。装置：5V输入，DC-DC输出3.3V，40MHz数字活动，板约10cm×7cm。测量：3m暗室、天线高1m、水平极化、转台旋转一周。此处原厂未完整给出磁芯型号/材料/数量/匝数/位置尺寸、RBW、检波器和限值，不能补写；另一段150MHz细化分析的10kHz RBW不移用。');
}
// 10. A native evidence table develops interpretation rather than adding harder physics.
{
 const s=await slide('怎样理解这个案例的改善？','依据 [5] 原厂文字与频谱整理；下半部分为本组对结果的分析');
 table(s,[['证据或问题','可以得到的判断'],['同一装置加装前后频谱不同','该装置的线缆磁芯措施对部分辐射峰值有效，原厂报告超过10 dB'],['加装后仍有辐射','噪声可能来自其他路径；单独处理线缆并不覆盖所有板级噪声'],['换一个磁环或另一台设备','材料、频段、路径和安装改变后，需要重新比较效果'],['能否据此认定整机合规','还需适用标准、完整频段、工况与检波等信息，本图不能单独证明']],[350,800],65,157,1150,283,22);
 heading(s,'本组判断：磁环应针对明确的干扰路径',65,472,1148);
 para(s,'磁环的价值在于处理特定线缆上的目标干扰。先确定噪声频段和电流模式，再结合器件曲线选候选磁环，最后在一致条件下比较加装前后结果。这样才能区分器件阻抗、单个峰值改善和整机表现。',65,522,1148,82);
 text(s,'本次工作是资料调研与结果分析，测量数据来自原厂。',65,619,1148,37,24,C.red,true);
 notes(s,60,'这页不再增加新公式，而是练习判断证据。一个真实案例说明在那台装置上有效，不能直接推出所有设备都能有相同改善，也不能据图认定达标。我们的小组分析是把器件曲线与实际回路联系起来，再坚持同条件前后比较。',citations.r5+'\n'+citations.r4,'dB是前后电平差的表达；本案例的绝对辐射电平单位为dBμV/m。它不同于LISN测得的dBμV传导电压，也不同于器件阻抗Ω。');
}
// 11. Three plain-language conclusions suitable for a single undergraduate presenter.
{
 const s=await slide('调研结论与使用判断','综合 [1]–[5]；围绕二线直流供电的共模干扰解释');
 heading(s,'1. 先看电流方向，理解磁环对什么起作用',65,156,1148);
 para(s,'去回两根供电线一起穿过磁环，正常工作电流的磁化作用大致抵消，共模分量的作用相加。这是整体穿芯可以针对共模干扰的基础。',65,206,1148,68);
 rule(s,65,295,1150);
 heading(s,'2. 按目标频段看曲线，匝数和位置都要匹配',65,316,1148);
 para(s,'磁环的阻抗随频率变化。增加匝数可以改善某些频段，也会改变响应。选型和安装应围绕实际噪声路径展开，不能仅凭外形或匝数判断效果。',65,366,1148,68);
 rule(s,65,454,1150);
 heading(s,'3. 用前后对比评价效果，保留使用边界',65,475,1148);
 para(s,'村田案例说明线缆磁芯措施可降低部分辐射峰值，但其他噪声路径仍可能存在。器件阻抗提供选型线索，实际改善需要具体装置中的测量验证。',65,525,1148,68);
 text(s,'磁环是针对线缆干扰路径的工程措施，选对、接对后再验证。',65,621,1148,36,24,C.red,true);
 notes(s,45,'最后归纳三句话：先看电流方向，再按频率和路径选用，最后用前后对比验证。磁环很常见，但不能只凭看起来更大或绕得更多来判断效果。谢谢老师和同学。',Object.values(citations).join('\n'));
}
// 12. Full primary-source bibliography; deeper Q&A stays in notes rather than extending the lecture.
{
 const s=await slide('参考资料与答疑','引用均为原厂公开资料；访问日期2026-10-20，完整链接见备注与配套讲稿');
 table(s,[['编号','原始资料','本次用途'],['[1]','Würth Elektronik. Cable Solutions\n原厂 Product Navigator 页面','线缆磁环的用途、材料频段\n与安装背景'],['[2]','Murata. Conductor conduction and common mode\nSection 1, Chapter 5，Fig.5-2-3／5-2-10','两种电流模式、外部回路\n及线缆穿磁环的原图'],['[3]','Fair-Rite. Catalog, 17th Edition\nTechnical Information，印刷页144、152','阻抗分量与频率；\n同磁环1／2／3匝曲线'],['[4]','Fair-Rite. Round Cable Snap-It 0431164281\n原厂产品数据页，单匝典型阻抗表','6个频点的典型阻抗\n比较同一器件的频率响应'],['[5]','Murata. Noise Suppression Measures\nat DC Power Cable Connection，Fig.2-14左图','试验板、测量条件与\n加磁环前后辐射频谱']],[86,706,358],65,157,1150,433,20.5);
 text(s,'答疑重点：正常供电为何受影响较小？为什么匝数并非越多越好？',65,621,1150,39,23,C.blue,true);
 notes(s,20,'这些资料均可以从原厂公开页面查到。课堂上重点已经讲完，后面如果老师提问，可以回到电流方向图、匝数曲线和案例结果解释。',Object.values(citations).join('\n\n'),'问1：为什么正常供电也穿过磁环却影响较小？答：去回电流的磁化作用大致抵消，共模分量则相加。\n问2：为什么有的线没有磁环也能使用？答：磁环只是EMC措施之一，设备也可能靠内部滤波、布局、线缆或屏蔽满足需要；正常运行本身不等于已证明EMC合规。\n问3：加磁环一定有改善吗？答：若目标干扰并非该线上的共模分量，或材料、频段、接法不合适，效果可能有限，应同条件比较。\n问4：是不是匝数越多越好？答：匝数会同时改变阻抗和频率响应，原厂曲线显示高频区可能交叉，所以按目标频段判断。');
}

await fs.mkdir(buildDir,{recursive:true});await fs.mkdir(`${workspaceDir}/输出`,{recursive:true});
const candidatePath=`${buildDir}/candidate.pptx`;
await(await PresentationFile.exportPptx(p)).save(candidatePath);
await fs.writeFile(`${buildDir}/notes.json`,JSON.stringify(slides.map((s,i)=>({slide:i+1,notes:s.speakerNotes.text})),null,2));
const nativeTableOwners=[8,10,12];
const result=await finalizePresentation({workspaceDir,candidatePath,finalPath,pythonExecutable:RUNTIME_PYTHON,explicitTotalSlideCount:12,integrityValidatorPath:`${SKILL_DIR}/container_tools/inspect_presentation_package_integrity.py`,layoutValidatorPath:`${SKILL_DIR}/container_tools/inspect_presentation_layout_geometry.py`,layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...nativeTableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],requiredNativeTableOwnerSlides:nativeTableOwners,requiredNativeChartOwnerSlides:[6],fontPolicy,materializeLiteralChartWorkbooks:true,verifyArtifactToolImport:true,receiptPath:`${buildDir}/${finalName}.validation.json`});
console.log(JSON.stringify({finalPath:result.finalPath,slideCount:result.packageIntegrity.slide_count,bytes:result.byteCount,sha256:result.finalSha256,package:result.packageIntegrity.status,layoutFindings:result.presentationLayout.finding_count,chartValidation:result.nativeChartValidation.passed}));
