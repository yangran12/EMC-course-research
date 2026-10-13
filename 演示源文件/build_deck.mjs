import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const workspaceDir = 'D:/PycharmProjects/电磁兼容';
const SKILL_DIR = 'C:/Users/Lenovo/.codex/plugins/cache/openai-primary-runtime/presentations/26.1011.11113/skills/presentations';
const RUNTIME_ROOT = 'C:/Users/Lenovo/.cache/codex-runtimes/codex-primary-runtime/dependencies';
const RUNTIME_PYTHON = `${RUNTIME_ROOT}/python/python.exe`;
process.env.RUNTIME_NODE_MODULES = `${RUNTIME_ROOT}/node/node_modules`;
const requireRuntime = createRequire(`${RUNTIME_ROOT}/node/node_modules/__runtime__.cjs`);
const { Presentation, PresentationFile } = await import(pathToFileURL(requireRuntime.resolve('@oai/artifact-tool')).href);
const { finalizePresentation, resolvePresentationFont, applyPresentationChartFont } = await import(pathToFileURL(`${SKILL_DIR}/container_tools/artifact_tool_utils.mjs`).href);
const font = resolvePresentationFont({ fontFamily: 'Microsoft YaHei', availableFonts: ['Microsoft YaHei'] });
const fontPolicy = { basis: 'design', families: [font] };
const assetDir = `${workspaceDir}/演示源文件/素材`;
const buildDir = `${workspaceDir}/.build/ppt`;
const finalName = process.argv[2] || '南师大风格_压电谐振与EMI滤波器小型化_定稿.pptx';
if(path.basename(finalName)!==finalName || !finalName.endsWith('.pptx')) throw new Error('输出参数须为新建 PPTX 文件名');
const finalPath = `${workspaceDir}/输出/${finalName}`;
const C = { blue:'#002060', accent:'#133FAD', red:'#C00000', text:'#222B3A', muted:'#5B6473', line:'#CBD3DF', pale:'#ECF4F7', white:'#FFFFFF' };
const p = Presentation.create({ slideSize: { width:1280, height:720 } });
const slides = [];

function text(s, v, x, y, w, h, size=28, color=C.text, bold=false, align='left') {
  const sh=s.shapes.add({ geometry:'textbox', position:{left:x,top:y,width:w,height:h}, fill:'none', line:{fill:'none',width:0} });
  sh.text=v;
  sh.text.style={typeface:font,fontSize:size,color,bold,alignment:align,verticalAlignment:'top',wrap:'square',autoFit:'none',insets:0};
  return sh;
}
function rule(s,x,y,w,color=C.line,width=1) { s.shapes.add({geometry:'line',position:{left:x,top:y,width:w,height:0},fill:'none',line:{style:'solid',fill:color,width}}); }
async function img(s,file,x,y,w,h,alt) {
  const b=await fs.readFile(`${assetDir}/${file}`);
  s.images.add({blob:b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength),contentType:'image/png',alt,fit:'contain',position:{left:x,top:y,width:w,height:h}});
}
async function slide(title, source='', backup=false) {
  const s=p.slides.add(); slides.push(s); s.background.fill=C.white;
  text(s,title,60,43,860,59,43,C.blue,true);
  await img(s,'南师大校标.png',956,28,264,50.66,'南京师范大学校标，来自用户提供参考稿');
  rule(s,60,118,1160,C.accent,2);
  text(s,source,60,668,1048,32,16,C.muted);
  text(s,backup?`备用 ${slides.length-14}`:`${String(slides.length).padStart(2,'0')}`,1130,668,90,32,16,C.muted,false,'right');
  return s;
}
function notes(s,time,talk,sources) {
  s.speakerNotes.text=`【建议讲解 ${time}】\n${talk}\n\n【资料来源】\n${sources}\n\n【衔接】${time==='备用'?'根据提问选择展开。':'讲清当前页后自然转到下一页，不逐字读正文。'}`;
}
function table(s, values, widths, x=60,y=170,w=1160,h=400,size=25) {
  const t=s.tables.add({rows:values.length,columns:values[0].length,left:x,top:y,width:w,height:h,columnWidths:widths,values});
  t.borders.assign({style:'solid',fill:C.line,width:1});
  for(let r=0;r<values.length;r++)for(let c=0;c<values[0].length;c++){
    const cell=t.getCell(r,c); cell.fill=r===0?C.blue:(r%2?C.white:'#F4F7FA');
    cell.text.style={typeface:font,fontSize:size,color:r===0?C.white:C.text,bold:r===0,verticalAlignment:'middle',wrap:'square',autoFit:'none'};
    cell.text.insets={left:16,right:12,top:10,bottom:10};
  }
  return t;
}
function chartFont(ch){ applyPresentationChartFont(ch,{fontFamily:font}); return ch; }
const R={
  r1:'Xu S, Xu S, Xu D, Qian Q, Sun W, Zhu J. A review on recent effort of conductive EMI suppression methods in high-frequency power converters. IET Power Electronics, 2022, 15(16):1921–1935. DOI:10.1049/pel2.12348. https://ietresearch.onlinelibrary.wiley.com/doi/10.1049/pel2.12348',
  r2:'Pei Y, Yan W, Ma H, Zhou M, Yang J. Study of High Frequency Characteristics Modeling and EMI Suppression of Common Mode Chokes. ACES Journal, 2023, 38(2):91–99. DOI:10.13052/2023.ACES.J.380203. https://journals.riverpublishers.com/index.php/ACES/article/view/18433',
  r3:'Hubert et al. Piezoelectric EMI Filter for Switched-Mode Power Supplies. IEEE TPEL, 2021, 36(6):6624–6643. DOI:10.1109/TPEL.2020.3043091. 原始摘要:https://cris.fau.de/publications/248817792/',
  r4:'Hubert F, Wich M, Dürbaum T, Rupitsch SJ. Reducing EMI Filter Size and Losses with a Novel Piezoelectric Interference-Suppression Component. APEC 2024:1506–1511. DOI:10.1109/APEC48139.2024.10509072. https://cris.fau.de/publications/322825671/?lang=de_DE',
  r5:'Zhang T, Yan W, Zhou M. Modeling and performance evaluation of filtering components for switch-mode power supplies based on piezoelectric effect. JESTCH, 2025, 66:102060. DOI:10.1016/j.jestch.2025.102060. https://www.sciencedirect.com/science/article/pii/S2215098625001156',
  r6:'Zhang T, Yan W, Zhou M. Modeling and Measurement of Conducted EMI Suppression Elements for Switch-Mode Power Supplies Based on the Inverse Piezoelectric Effect. IEEE TIM, 2025, 74:1–11. DOI:10.1109/TIM.2025.3557100.',
  r7:'Zhang T, Yan W, Zhou M. Conducted interference suppression for Switched-Mode power supplies based on an impedance resonance design of piezoelectric ceramics. AEU, 2025,190:155622. DOI:10.1016/j.aeue.2024.155622.',
  r8:'Boles JD et al. Opportunities, Progress, and Challenges in Piezoelectric-Based Power Electronics. IEEJ Journal of Industry Applications, 2023,12(3):254–263. DOI:10.1541/ieejjia.22008217. 作者公开稿:https://per.mit.edu/wp-content/uploads/2023/10/IEEJ_PiezoReview2022_advpub_22008217.pdf',
  r9:'Hubert F, Dürbaum T, Rupitsch S. Electromagnetic interference suppression components. EP3591819A1,公开日2020-01-08. https://patents.google.com/patent/EP3591819A1/en',
  r10:'Texas Instruments. Introduction to EMI in power supply design,教学资料,第17/19页. https://www.ti.com/content/dam/videos/external-videos/en-us/3/3816841626001/5839978751001.mp4/subassets/intoduction_to_emi_in_power_supply_design_final.pdf'
};

// 01 Cover. School artwork comes from the supplied reference without redrawing.
{
 const s=p.slides.add(); slides.push(s); s.background.fill=C.white;
 await img(s,'南师大校标.png',345,70,590,113.2,'南京师范大学组合校标');
 text(s,'开关电源 EMI 滤波器小型化\n与压电谐振抑制技术',100,268,1080,164,56,C.blue,true,'center');
 text(s,'电磁兼容课程文献调研',160,467,960,46,30,C.accent,false,'center');
 text(s,'汇报人：杨冉    授课老师：周孟夏',160,542,960,45,27,C.blue,false,'center');
 notes(s,'20 秒','老师、同学们好。本次汇报关注开关电源中的 EMI 滤波器。我们想回答：压电谐振器件怎样帮助缩小滤波器，这种帮助又在什么条件下成立。我们采用文献调研，并做了简化模型的计算分析。正常讲解到第14页，后面的页面供答疑使用。','校标与风格参考用户提供的中期答辩PPT。当前汇报姓名参考其封面；课程老师来自用户说明。');
}
// 02 Scope and research questions.
{
 const s=await slide('研究问题与范围','依据 [1]–[5] 的研究范围整理');
 text(s,'电源越来越紧凑，滤波器怎样跟着缩小？',65,165,1140,63,38,C.blue,true);
 table(s,[['研究对象','本次关注的内容'],['隔离型开关电源','以反激变换器案例为主'],['共模传导干扰','噪声回路、阻抗与滤波支路'],['压电谐振器件','抑制效果、小型化来源及适用条件']],[290,870],60,271,1160,260,28);
 text(s,'判断依据：干扰抑制效果与滤波器代价，需要同时考察。',65,577,1135,49,30,C.red,true);
 notes(s,'40 秒','本次范围比较窄：选择隔离型开关电源，主要看反激案例里的共模传导干扰。传统共模电感与抑制电容是比较基线。我们将先解释路径和模型，再看两项小型化结果。体积变小只是一个指标，是否适用于其他工况还需要测量条件和可靠性证据。',R.r1+'\n'+R.r3+'\n'+R.r5);
}
// 03 Actual source circuit plus qualitative noise equation.
{
 const s=await slide('开关电源中的共模噪声源','图：[9] Fig.3a，专利电路示意    原理：[1]');
 await img(s,'patent_flyback_baseline.png',60,165,694,431,'专利传统反激电源及输入滤波器示意');
 text(s,'寄生电容中的位移电流',807,185,405,52,31,C.blue,true);
 text(s,'i = Cₚ · dv/dt',807,263,405,70,40,C.red,true);
 text(s,'开关节点快速变化\n\n寄生电容耦合噪声\n\n噪声沿高频回路返回',807,365,405,220,29);
 text(s,'Cₚ 是未画出的寄生耦合电容，与原图抑制电容 Cᵧ 区分。',65,610,1128,35,22,C.muted);
 notes(s,'60 秒','左图是公开专利中的传统反激电源示意，用来定位一次侧开关节点和隔离变压器，不把它当作实测结果。快速电压变化会通过寄生电容产生位移电流。这里公式只作定性解释，实际电流还受完整回路的阻抗影响。一次与二次绕组之间、设备与参考平面之间都可能存在寄生耦合。',R.r1+'\n'+R.r9+'，PDF第11页Fig.3a。');
}
// 04 CM equivalent loop. The source uses a common measurement plane, not merged converter grounds.
{
 const s=await slide('共模传导干扰的回路','图：[10] 第19页，共模教学等效模型');
 await img(s,'ti_common_mode_loop.png',60,170,705,428,'TI共模等效回路：LISN、共模电感、开关噪声源和寄生电容');
 text(s,'VSW：开关噪声源',820,191,380,48,29,C.blue,true);
 text(s,'CPAR：对参考平面的寄生电容\n\nL1、L2：两条电源线\n共模电流具有同向分量\n\n50 Ω：LISN 测量端终接',820,272,389,275,25);
 text(s,'图中 GND 代表模型的共同参考平面',70,612,1120,39,23,C.red);
 notes(s,'70 秒','沿左图从VSW出发：噪声通过CPAR耦合到共同参考平面，再经LISN的两条50欧路径返回输入导线，并回到噪声源，构成高频回路。共模电流在两条输入线上具有同向分量。这个TI图是共模教学等效模型，两个GND表示同一测量参考平面。实际隔离电源的一次侧返回点、二次侧返回点、机壳与PE必须分别识别。压电支路的效果要放回真实回路讨论。',R.r10+'\n'+R.r1);
}
// 05 Baseline, high frequency parasitics.
{
 const s=await slide('传统滤波器的体积与高频约束','图：[2] Fig.2    [2] 建模频率范围 9 kHz–200 MHz');
 await img(s,'aces_filter.png',60,169,716,333,'ACES原文EMI滤波器实物和电路');
 text(s,'共模电感',824,178,376,52,32,C.blue,true);
 text(s,'提供共模串联阻抗\n体积与绕组、磁芯有关',824,238,376,110,27);
 text(s,'抑制电容与寄生参数',824,373,376,52,31,C.blue,true);
 text(s,'改变高频噪声路径\n真实阻抗随频率变化',824,433,376,110,27);
 text(s,'更大电感值未必在整个频段提供更强抑制。',67,585,1135,49,31,C.red,true);
 notes(s,'65 秒','这篇南师大团队论文为传统滤波器建立高频模型。原图同时给出器件实物与电路。共模电感通过串联阻抗抑制共模电流，电容提供其他高频路径。绕组寄生电容、漏感和磁芯特性使实际器件偏离理想模型。因而小型化不能只比较标称电感和体积，需要看测量频段中的阻抗及实际滤波效果。该文模型频率范围9千赫到200兆赫，模型范围不能等同于某一产品的合规测试范围。',R.r2+'，PDF第3页（期刊第93页）Fig.2。');
}
// 06 Actual piezo branch positions, as a patent principle illustration.
{
 const s=await slide('压电谐振支路的接入位置','图：[9] Fig.3b，专利原理示意，非论文实测');
 await img(s,'patent_piezo_filter.png',60,166,698,432,'专利示意中压电元件标记10及其在一次二次侧之间的支路位置');
 text(s,'谐振频率附近',803,177,404,53,31,C.blue,true);
 text(s,'电机械耦合形成\n选频低阻抗通道',803,239,404,101,28);
 text(s,'其他频率',803,380,404,52,31,C.blue,true);
 text(s,'可表现出电容性行为',803,442,404,58,28);
 text(s,'跨一次、二次侧接入\n需核对隔离与器件要求',803,533,404,93,26,C.red);
 notes(s,'60 秒','与第3页传统电路对照，专利图中标记10是压电抑制元件。图右侧一次二次侧之间的支路对应传统Cy位置。这是概念性接入示意，不能把其中所有地符号当成同一个节点，也不能据它声称实测达标。压电材料在电压作用下发生机械形变，机械响应再影响端口电流，于是谐振附近出现低阻抗。其他频率可呈电容性行为。实际采用时，要确认回路、频率匹配和隔离资格，不能直接把普通陶瓷片视为已有安全认证的Y电容。',R.r9+'，PDF第11页Fig.3b，原文[0039]说明图用于概念性示意。\n'+R.r3+'\n'+R.r8);
}
// 07 BVD single-mode model.
{
 const s=await slide('压电器件的 BVD 等效模型','图：[8] Fig.2    原图 Cp 对应本文 C₀，L/C/R 对应 Lm/Cm/Rm');
 await img(s,'bvd_original.png',138,189,294,308,'BVD模型，静态电容与串联RLC运动支路并联');
 table(s,[['参数','物理含义'],['C₀','静态端口电容'],['Lₘ、Cₘ','等效惯性与弹性'],['Rₘ','等效机械及电气损耗']],[175,515],530,184,690,303,27);
 text(s,'C₀ 与串联 Rₘ–Lₘ–Cₘ 支路并联',86,556,1108,54,34,C.red,true,'center');
 notes(s,'85 秒','BVD是Butterworth–Van Dyke模型。原图标Cp，这里记成C0。左边静态电容与右边串联RLC运动支路并联。运动支路的电感和电容描述等效惯性与弹性，电阻描述损耗；它们是等效参数，并不意味着实物中额外串接三个分立元件。单模模型主要用于某个振动模态附近。多个模态需要多个运动支路，不能用单模参数覆盖全部高频行为。',R.r8+'，作者公开稿PDF第2页Fig.2。');
}
// 08 Reproducible native data chart, illustrative parameter set.
const C0=1e-9, Cm=50e-12, Rm=20, fsHz=300e3, Lm=1/((2*Math.PI*fsHz)**2*Cm);
const fpHz=fsHz*Math.sqrt(1+Cm/C0);
function impedance(f){const w=2*Math.PI*f;const im=w*Lm-1/(w*Cm);const d=Rm*Rm+im*im;const yr=Rm/d,yi=w*C0-im/d;return 1/Math.hypot(yr,yi);}
const round6=n=>Number(n.toFixed(6));
{
 const s=await slide('谐振处的阻抗低谷与反谐振峰','模型计算示例，非实测数据    模型依据 [8]');
 const freq=[...Array.from({length:145},(_,i)=>286+i*0.25),round6(fpHz/1000)].sort((a,b)=>a-b);
 chartFont(s.charts.add('scatter',{position:{left:60,top:166,width:785,height:421},categories:freq.map(String),series:[{name:'单模 BVD',xValues:freq,values:freq.map(f=>round6(20*Math.log10(impedance(f*1000)))),line:{fill:C.accent,width:3,style:'solid'},marker:{symbol:'none'}}],scatterOptions:{style:'line'},hasLegend:false,xAxis:{title:'频率 / kHz',min:286,max:322,majorUnit:6,textStyle:{fontSize:19},majorGridlines:null},yAxis:{title:'20 log₁₀(|Z| / 1 Ω)',min:20,max:110,majorUnit:20,textStyle:{fontSize:19},majorGridlines:{fill:'#DDE3EA',width:1}},chartFill:'none',plotAreaFill:'none'}));
 text(s,'串联谐振',879,191,331,50,30,C.blue,true);
 text(s,'fₛ = 300 kHz\n阻抗接近低谷',879,250,331,105,27);
 text(s,'反谐振',879,385,331,49,30,C.red,true);
 text(s,'fₚ ≈ 307.4 kHz\n出现高阻抗峰',879,443,331,107,27);
 text(s,'示例参数：C₀ = 1 nF，Cₘ = 50 pF，Rₘ = 20 Ω，Lₘ ≈ 5.63 mH',65,605,1150,35,21,C.muted);
 notes(s,'60 秒','这张图由单模BVD公式计算，参数是教学选择，没有对应特定陶瓷器件。横轴为千赫，纵轴是阻抗幅值的分贝表达，参考值为1欧，不是传导骚扰电压。300千赫附近串联支路发生谐振，端口阻抗出现低谷；约307.4千赫是低损耗近似得到的反谐振位置，阻抗升高。低阻抗支路是否减小LISN端噪声，仍取决于它接在哪条路径上。',R.r8+'\n本组计算：Z=[jωC0+1/(Rm+jωLm+1/(jωCm))]^-1，fs=300kHz，C0=1nF，Cm=50pF，Rm=20Ω，Lm=1/((2πfs)^2Cm)。数据可由本演示源文件复算。');
}
// 09 Frequency drift tradeoff, motional branch only.
{
 const s=await slide('谐振匹配与频偏敏感性','运动支路的理想计算示例，非 EMI 衰减实测    频移讨论见 [3]');
 const det=Array.from({length:81},(_,i)=>round6(-1+i*0.025));
 const ch=s.charts.add('scatter',{position:{left:60,top:163,width:754,height:416},categories:det.map(String),series:[50,200].map((q,i)=>({name:`Q = ${q}`,xValues:det,values:det.map(d=>{const r=1+d/100;return round6(Math.sqrt(1+q*q*(r-1/r)**2))}),line:{fill:i?C.red:C.accent,width:3,style:'solid'},marker:{symbol:'none'}})),scatterOptions:{style:'line'},hasLegend:true,legend:{position:'bottom',textStyle:{fontSize:21}},xAxis:{title:'频率偏差 / %',min:-1,max:1,majorUnit:0.5,textStyle:{fontSize:19},majorGridlines:null},yAxis:{title:'|Zₘ| / Rₘ',min:0,max:5,majorUnit:1,crossesAt:-1,textStyle:{fontSize:19},majorGridlines:{fill:'#DDE3EA',width:1}},chartFill:'none',plotAreaFill:'none'}); chartFont(ch);
 text(s,'更高 Q 值',846,189,352,53,32,C.blue,true);
 text(s,'谐振更尖锐\n频偏时相对谷值上升更快',846,258,352,132,28);
 text(s,'应用判断',846,420,352,51,31,C.blue,true);
 text(s,'核对温度与开关频率变化\n多模态可匹配多个离散峰',846,482,352,113,27);
 text(s,'模型显示的是支路阻抗变化，实际抑制量还受 C₀ 与外部回路影响。',64,614,1150,37,22,C.muted);
 notes(s,'80 秒','这页是本组的计算分析。只取串联运动支路，比较两个人为选定的Q值。横轴是相对串联谐振频率的偏差，纵轴是支路阻抗相对电阻Rm的比值。Q=200的曲线更陡，说明峰值匹配对频移更敏感。曲线不能直接换算成EMI改善量，因为完整BVD还有C0，系统还有其他阻抗。文献讨论了器件频移与开关频率偏移；多模态设计能针对多个离散峰，但宽频适用性仍需单独验证。',R.r3+'\n'+R.r5+'\n计算公式：|Zm|/Rm=sqrt(1+Q²(r−1/r)²)，r=f/fs，Q=ωsLm/Rm。本页仅作频偏敏感性的教学比较。');
}
// 10 Case A literal values from primary abstract, local baseline only.
function caseChart(s,categories,values,format='0%'){
 const ch=s.charts.add('bar',{position:{left:65,top:233,width:733,height:344},categories,series:[{name:'该论文基线',values:categories.map(()=>1),fill:'#A7B4C9'},{name:'采用压电方案',values,fill:C.accent}],barOptions:{direction:'column',grouping:'clustered',gapWidth:95},hasLegend:true,legend:{position:'bottom',textStyle:{fontSize:21}},xAxis:{textStyle:{fontSize:22},majorGridlines:null},yAxis:{title:'相对该论文基线',min:0,max:1.1,majorUnit:.25,numberFormatCode:'0%',textStyle:{fontSize:19},majorGridlines:{fill:'#DDE3EA',width:1}},dataLabels:{showValue:true,position:'outEnd',numberFormatCode:format,textStyle:{fontSize:21}},chartFill:'none',plotAreaFill:'none'}); chartFont(ch); }
{
 const s=await slide('案例 A：滤波器体积与直流损耗','数据：[4] APEC 2024，作者摘要报告    图中各项基线归一为 100%');
 text(s,'Hubert 等，反激变换器应用',66,166,1150,54,34,C.blue,true);
 caseChart(s,['滤波器体积','滤波器 DC 损耗'],[.36,.46]);
 text(s,'体积减少 64%',849,259,365,65,36,C.red,true);
 text(s,'DC 损耗减少 54%',849,361,365,67,34,C.red,true);
 text(s,'统计对象是滤波器\n整机尺寸与效率需另行评估',849,476,365,112,25);
 notes(s,'65 秒','APEC2024案例摘要报告滤波器体积减少64%，滤波器直流损耗减少54%。图把该论文自己的基线分别记成100%，采用压电方案后对应剩余36%和46%。DC损耗指滤波器的直流损耗，不能说整机损耗减少54%，也不能直接推成整机效率提高54%。这里依据作者原始摘要；功率、具体器件配置与完整检测条件还没有通过全文确认，所以结论保留在该反激案例。',R.r4+'\n图表为对摘要数值的算术重表达：剩余比例=1−减少比例。');
}
// 11 Case B.
{
 const s=await slide('案例 B：无源滤波器体积与重量','数据：[5] JESTCH 2025，作者摘要报告    图中各项基线归一为 100%');
 text(s,'Zhang、Yan、Zhou，反激变换器应用',66,166,1150,54,33,C.blue,true);
 caseChart(s,['无源滤波器体积','无源滤波器重量'],[.0929,.1128],'0.00%');
 text(s,'体积减少 90.71%',836,259,382,65,35,C.red,true);
 text(s,'重量减少 88.72%',836,361,382,67,35,C.red,true);
 text(s,'几何尺寸与多模态谐振设计\n是论文研究的重要内容',836,476,382,112,25);
 notes(s,'65 秒','第二项是2025年JESTCH论文，作者为Zhang、Yan、Zhou。出版者摘要报告反激案例中无源滤波器体积减少90.71%，重量减少88.72%。本页展示剩余9.29%和11.28%。原摘要还强调压电器件模型以及几何尺寸与多个谐振频率的关系。全文尚未取得，所以我们不补写未确认的功率或频率，更不把该结果直接与上一页排序。',R.r5+'\n原始出版者页面及DOAJ元数据共同核对摘要与书目。图为摘要数据算术重表达。');
}

// 12 Evidence comparability table.
{
 const s=await slide('文献结果的可比条件','依据 [4]、[5] 原始摘要整理');
 table(s,[['对照维度','案例 A：APEC 2024','案例 B：JESTCH 2025'],['统计对象','滤波器体积、DC 损耗','无源滤波器体积、重量'],['共同场景','反激变换器应用','反激变换器应用'],['基线与工况','各自论文的传统方案\n具体参数待全文核对','各自论文的无源方案\n具体参数待全文核对'],['证据可支持','各自案例具有小型化潜力','各自案例具有小型化潜力']],[230,465,465],60,170,1160,365,24);
 text(s,'比较优劣还需统一功率、频率、滤波器边界及测量条件。',65,579,1145,56,31,C.red,true);
 notes(s,'65 秒','两篇文章都有反激案例，但统计量不同。案例A含直流损耗，B含重量，体积也使用各自的设计基线。仅凭64%与90.71%无法确定哪种方案在同等条件下更好。跨论文比较需要统一电源功率、输入输出工况、开关频率、滤波器组成与体积计算，以及检测频段、检波器和限值。没有确认的字段保留未确认，比补写参数更可靠。',R.r4+'\n'+R.r5);
}
// 13 Group analysis explicitly separated from reported results.
{
 const s=await slide('本组分析：压电方案的适用条件','本组基于 [3]、[5]、[7] 的机制与案例分析');
 table(s,[['条件','对方案选择的影响'],['少数稳定的干扰峰','更适合通过谐振针对性抑制'],['峰值漂移或宽频噪声','需覆盖频偏，可能仍需常规滤波'],['接入路径明确','低阻抗支路才可能减少测量端噪声'],['隔离与器件资格','接入一次、二次侧或 PE 前需核对要求'],['安装、温度与长期使用','器件参数与机械边界应纳入验证']],[305,855],60,165,1160,429,26);
 notes(s,'75 秒','这是本组的分析，不是论文额外报告的数据。压电谐振特别适合需要重点压低少数稳定峰值的场景。如果频率随控制策略或温度明显变化，就必须核对实际有效带宽。接入位置同样关键：低阻抗支路如果改变了错误的返回路径，未必减小LISN端噪声。跨隔离边界使用还需要核对绝缘、额定电压及器件资格。安装约束会影响机械振动，长期参数稳定性也关系到推广。',R.r3+'\n'+R.r5+'\n'+R.r7);
}
// 14 Close the main 14-minute narrative.
{
 const s=await slide('调研结论','主要依据 [2]–[5]、[8]、[9]    参考书目及链接见备用页备注');
 text(s,'压电谐振为滤波器小型化提供了一条选频途径。',65,180,1140,71,38,C.blue,true);
 rule(s,65,280,1140);
 text(s,'机制',66,329,175,56,32,C.accent,true);
 text(s,'通过电机械谐振改变端口阻抗及噪声返回路径',245,327,964,65,31);
 text(s,'证据',66,425,175,56,32,C.accent,true);
 text(s,'案例摘要报告特定反激滤波器的体积明显缩减',245,423,964,65,31);
 text(s,'边界',66,521,175,56,32,C.red,true);
 text(s,'推广取决于频率匹配、测量工况和器件使用条件',245,519,964,65,31);
 notes(s,'30 秒','我们归纳三点。压电谐振通过改变端口阻抗和高频返回路径，提供针对某些噪声峰值的抑制机会。案例摘要已经报告明显的滤波器缩减，但这些数字各自依赖其比较基线。进一步应用时，必须把频率匹配、测量条件与器件资格放在一起看。谢谢老师和同学，后面是供答疑展开的资料。',R.r2+'\n'+R.r3+'\n'+R.r4+'\n'+R.r5+'\n'+R.r8+'\n'+R.r9);
}
// Backup 15–19.
{
 const s=await slide('参考文献：核心研究','编号与正文一致',true);
 const rows=[['编号','文献与出版信息'],['[1]','Xu et al. 传导 EMI 抑制综述\nIET Power Electronics, 2022. DOI:10.1049/pel2.12348'],['[2]','Pei et al. 共模电感高频建模\nACES, 2023. DOI:10.13052/2023.ACES.J.380203'],['[3]','Hubert et al. Piezoelectric EMI Filter\nIEEE TPEL, 2021. DOI:10.1109/TPEL.2020.3043091'],['[4]','Hubert et al. Reducing EMI Filter Size and Losses\nAPEC, 2024. DOI:10.1109/APEC48139.2024.10509072'],['[5]','Zhang, Yan, Zhou. 压电滤波建模与性能评价\nJESTCH, 2025. DOI:10.1016/j.jestch.2025.102060']];
 table(s,rows,[115,1045],60,161,1160,453,22);
 notes(s,'备用','本页列核心文献。第2篇可读全文，是传统滤波基线。第3到第5篇的数值或机制摘要应保留其证据层级。完整英文题名和网址写在本页备注，方便课后查阅。',Object.values(R).slice(0,5).join('\n\n'));
}
{
 const s=await slide('参考文献：模型与扩展阅读','编号与正文一致',true);
 table(s,[['编号','文献与用途'],['[6]','Zhang, Yan, Zhou. 压电抑制元件建模与测量\nIEEE TIM, 2025. DOI:10.1109/TIM.2025.3557100'],['[7]','Zhang, Yan, Zhou. 压电陶瓷阻抗谐振设计\nAEU, 2025. DOI:10.1016/j.aeue.2024.155622'],['[8]','Boles et al. 压电功率电子综述与 BVD 模型\nIEEJ, 2023. DOI:10.1541/ieejjia.22008217'],['[9]','Hubert, Dürbaum, Rupitsch. EMI 抑制元件专利\nEP3591819A1，2020，原理图来源'],['[10]','Texas Instruments. Introduction to EMI in power supply design\n公开教学资料，共模回路与 LISN 解释']],[115,1045],60,161,1160,453,22);
 notes(s,'备用','第6、7篇是扩展阅读，与第5篇作者组相近，不能自动计为三个独立团队的验证。第8篇提供原始BVD和模态图。第9篇是公开专利，用于说明电路原理。第10项是器件厂商的教学材料，用于解释测量与共模路径。',Object.values(R).slice(5).join('\n\n'));
}
{
 const s=await slide('备用：模型公式与计算口径','理想单模 BVD，符号与正文一致',true);
 text(s,'Z(ω) = [jωC₀ + 1 / (Rₘ + jωLₘ + 1 / (jωCₘ))]⁻¹',65,177,1150,77,35,C.blue,true);
 text(s,'fₛ = 1 / (2π√(LₘCₘ))',65,296,1110,63,35);
 text(s,'fₚ ≈ fₛ√(1 + Cₘ/C₀)',65,392,1110,63,35);
 text(s,'Q = ωₛLₘ / Rₘ',65,488,1110,59,35);
 text(s,'反谐振频率使用低损耗近似。单模模型主要描述对应模态附近的特征。',65,594,1150,54,25,C.red);
 notes(s,'备用','如果老师问到公式，先说明静态电容和运动支路并联，所以总导纳是两部分相加。串联谐振条件是ωLm−1/(ωCm)=0。理想低损耗反谐振条件来自两支路导纳虚部相消。频偏图取的是运动支路，所以不能单独代表端口阻抗或系统EMI衰减。参数均为教学示例，未针对论文器件做拟合。',R.r8+'\n正文计算参数：C0=1nF、Cm=50pF、Rm=20Ω、fs=300kHz、Lm约5.62895mH、fp约307.4085kHz。');
}
{
 const s=await slide('备用：滤波器缩减与整机缩减','本组口径分析，数字为教学假设',true);
 text(s,'整机体积还包含功率级、控制、散热与外壳。',65,171,1150,64,35,C.blue,true);
 text(s,'设滤波器原占整机体积的比例为 α',65,292,1150,57,31);
 text(s,'r整机 = α · r滤波器   （r 表示体积缩减比例）',65,372,1150,64,35,C.red,true);
 text(s,'假设 α = 20%，滤波器缩减 64%\n则简化估算的整机体积缩减为 12.8%',65,470,1150,104,32);
 text(s,'适用假设：体积可加总，其他部分保持不变。实际外形还受布局约束。',65,604,1150,45,23,C.muted);
 notes(s,'备用','如果有人问滤波器减少64%是不是充电器就减少64%，用这页解释。假设原始体积可以加总，并且除了滤波器其他部分保持不变，滤波器占比alpha乘以其缩减比例，才得到整机体积缩减估算。本页20%是假设值，64%取自案例A用于举例，乘积12.8%并不是该论文实测的整机结果。实际外形尺寸还受布局、散热和外壳约束。',R.r4+'\n本组简化推导：Vnew=Vother+(1−rfilter)Vfilter。假设体积可加且Vother保持不变，因此rtotal=alpha*rfilter。');
}
{
 const s=await slide('备用：常见问题与回答要点','答疑依据正文机制、案例及其适用边界',true);
 table(s,[['问题','回答要点'],['是否能取代全部传统滤波？','需按噪声频谱和有效带宽判断，可能与常规滤波配合'],['压电器件接在什么位置？','依据真实共模回路，核对一次、二次侧与 PE 的区别'],['单峰衰减能否证明合规？','需规定频段、限值、工况及检波方式的完整测量'],['dB 与 dBμV 有何区别？','dB 表示比值或增益、衰减量\ndBμV 是相对 1 μV 的电压电平'],['主动滤波怎样比较？','还需考虑辅助供电、环路与带宽，本课题以压电方案为主']],[392,768],60,161,1160,453,24);
 notes(s,'备用','答疑时先给短结论，再回到对应图或模型。不能保证压电器件全面取代传统滤波。低阻抗取决于频率，抑制效果取决于路径。单个频点的衰减不等于全频段达标。传导骚扰电压通常用dBμV表达，两次电平相减得到的改善量用dB。主动EMI滤波可作其他路线讨论，但本次没有对其开展完整的跨方法性能评估。',R.r1+'\n'+R.r3+'\n'+R.r10);
}

{
 const s=await slide('备用：振动模态与多模设计','图：[8] Fig.12    多模设计讨论见 [5]、[7]',true);
 text(s,'不同几何形状与极化方向允许不同的振动模态。',63,166,1148,60,34,C.blue,true);
 await img(s,'vibration_modes_original.png',72,294,1136,180,'厚度伸缩、厚度剪切、轮廓伸缩和径向模态，Boles等Fig12');
 text(s,'厚度伸缩',78,500,265,45,27,C.accent,true,'center');
 text(s,'厚度剪切',359,500,265,45,27,C.accent,true,'center');
 text(s,'轮廓伸缩',641,500,265,45,27,C.accent,true,'center');
 text(s,'径向振动',923,500,265,45,27,C.accent,true,'center');
 text(s,'多模模型可并联多个运动支路，参数需来自器件建模或测量。',72,587,1136,50,27,C.red,true);
 notes(s,'备用','材料、尺寸、电极和安装条件决定能激发哪些模态。单模BVD只描述其中一个模态附近的行为。若器件具有多个可用模态，可在模型中增加运动支路，从而描述多个谐振特征。多个谐振频率并不自动形成连续宽频带抑制，仍需按实际噪声峰值和系统回路判断。',R.r8+'，作者公开稿PDF第5页Fig.12。\n'+R.r5+'\n'+R.r7);
}

await fs.mkdir(buildDir,{recursive:true});
await fs.mkdir(`${buildDir}/draft_previews`,{recursive:true});
await fs.mkdir(`${workspaceDir}/输出`,{recursive:true});
const candidatePath=`${buildDir}/candidate.pptx`;
await (await PresentationFile.exportPptx(p)).save(candidatePath);
await fs.writeFile(`${buildDir}/notes.json`,JSON.stringify(slides.map((s,i)=>({slide:i+1,notes:s.speakerNotes.text})),null,2));
const nativeTableOwners=[2,7,12,13,15,16,19];
const result=await finalizePresentation({workspaceDir,candidatePath,finalPath,pythonExecutable:RUNTIME_PYTHON,integrityValidatorPath:`${SKILL_DIR}/container_tools/inspect_presentation_package_integrity.py`,layoutValidatorPath:`${SKILL_DIR}/container_tools/inspect_presentation_layout_geometry.py`,layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...nativeTableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],requiredNativeTableOwnerSlides:nativeTableOwners,requiredNativeChartOwnerSlides:[8,9,10,11],fontPolicy,materializeLiteralChartWorkbooks:true,verifyArtifactToolImport:true,receiptPath:`${buildDir}/${finalName}.validation.json`});
console.log(JSON.stringify({finalPath:result.finalPath,slideCount:result.packageIntegrity.slide_count,bytes:result.byteCount,sha256:result.finalSha256,package:result.packageIntegrity.status,layoutFindings:result.presentationLayout.finding_count,chartValidation:result.nativeChartValidation.passed}));
