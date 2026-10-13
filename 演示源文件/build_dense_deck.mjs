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
const buildDir = `${workspaceDir}/.build/ppt_dense`;
const finalName = process.argv[2] || '南师大风格_压电谐振与EMI滤波器小型化_内容增强版.pptx';
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
    cell.text.insets={left:16,right:12,top:8,bottom:8};
  }
  t.rows[0].height=44;
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

// Academic layouts follow the supplied reference: paragraphs, source figures and data interpretation.
function heading(s,v,x,y,w=550,size=29,color=C.accent){return text(s,v,x,y,w,40,size,color,true);}
function para(s,v,x,y,w,h,size=26){
  const runs=v.split(/(\*\*.*?\*\*)/g).filter(Boolean).map(a=>a.startsWith('**')?{run:a.slice(2,-2),style:{bold:true,color:C.blue}}:{run:a});
  const sh=text(s,v.replaceAll('**',''),x,y,w,h,size);
  sh.text=[{runs:runs.map(a=>({...a,style:{typeface:font,fontSize:`${size}px`,color:C.text,bold:false,...a.style}}))}];
  return sh;
}
function emphasis(s,v,y=608){return text(s,v,65,y,1150,48,26,C.red,true);}
function frame(s,x,y,w,h){s.shapes.add({geometry:'rect',position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'dash',fill:C.accent,width:1.5}});}

// 01 Cover keeps the supplied reference's minimal opening layout.
{
 const s=p.slides.add();slides.push(s);s.background.fill=C.white;
 await img(s,'南师大校标.png',345,70,590,113.2,'南京师范大学组合校标');
 text(s,'开关电源 EMI 滤波器小型化\n与压电谐振抑制技术',100,268,1080,164,56,C.blue,true,'center');
 text(s,'电磁兼容课程文献调研',160,467,960,46,30,C.accent,false,'center');
 text(s,'汇报人：杨冉    授课老师：周孟夏',160,542,960,45,27,C.blue,false,'center');
 notes(s,'20 秒','老师、同学们好。本次调研关注开关电源中的EMI滤波器。我们想回答：压电谐振器件如何帮助缩小滤波器，这种帮助在什么条件下成立。正常汇报讲到第14页，之后的页面用于答疑。','校标与版式参考用户提供的中期答辩PPT；姓名来自参考稿，课程老师来自用户说明。');
}
// 02 Dense paired paragraphs and a compact native table, rather than a sparse overview.
{
 const s=await slide('研究背景与调研问题','依据 [1]–[5] 整理；讨论对象以隔离型反激电源的共模传导干扰为主');
 frame(s,60,153,557,283);frame(s,641,153,579,283);
 heading(s,'电源小型化中的滤波器约束',78,170,521);
 para(s,'高频开关有助于缩小功率级磁性元件，但快速电压变化也会经寄生电容激励共模噪声。传统滤波器依靠**共模电感与抑制电容**改善传导干扰，其体积、高频寄生参数和损耗都应纳入设计。因此，小型化应同时评价**干扰抑制效果与体积、损耗代价**。',78,221,521,204,25);
 heading(s,'压电谐振提供的研究切入点',659,170,543);
 para(s,'压电器件通过电机械耦合，在谐振频率附近形成低阻抗特征。放入合适的共模返回路径，有机会针对主要干扰峰改善滤波，并减少部分传统滤波器体积。本次沿**噪声路径—等效模型—文献案例—工程条件**展开，解释有效性及适用边界。',659,221,543,204,25);
 table(s,[['调研问题','需要核对的证据'],['怎样抑制噪声？','接入位置、端口阻抗、共模返回路径'],['为什么能缩小？','传统滤波基线、器件组成、体积和损耗口径'],['怎样评价推广？','频率变化、测量工况、隔离与安装条件']],[260,900],60,454,1160,172,24);
 notes(s,'40 秒','这页先交代研究矛盾。功率级做小之后，EMI滤波仍要占用空间。压电谐振的价值在于针对性改变阻抗，但小型化是否成立还需要确认路径和比较对象。表格的三项问题分别对应后面机制、案例和适用条件；汇报时只提三个关键词，不逐字读两段背景。',R.r1+'\n'+R.r2+'\n'+R.r3+'\n'+R.r5);
}
// 03 Circuit plus explanatory paragraphs and a conditional numerical example.
{
 const s=await slide('开关电源中的共模噪声源','图：[9] Fig.3a，专利原理示意；噪声机制依据 [1]、[10]');
 heading(s,'传统反激电源与输入滤波器',63,149,635);
 await img(s,'patent_flyback_baseline.png',60,193,650,325,'EP3591819A1 Fig3a传统反激电源及滤波器原理示意');
 para(s,'原图 Cᵧ 是有意设置的抑制电容；**Cₚ 表示未画出的寄生耦合电容**，两者不能混用。一次侧、二次侧和机壳应按真实连接分别识别。',65,540,639,92,25);
 heading(s,'快速电压变化激励位移电流',746,150,473);
 text(s,'iₚ = Cₚ · dv/dt',746,202,473,57,37,C.red,true);
 para(s,'开关节点 VSW 的电压边沿，经绕组间及对参考平面的寄生电容耦合出高频电流。电流沿输入导线、测量网络及其他耦合路径返回，构成**完整的共模回路**。',746,275,473,147,26);
 heading(s,'影响因素与数量级',746,436,473,27);
 para(s,'假设 Cₚ = 20 pF、dv/dt = 10 V/ns，理想电容关系给出瞬时电流 0.2 A。该数值仅说明数量级，实际电流还取决于边沿波形与回路阻抗。',746,482,473,118,25);
 notes(s,'60 秒','左图用于定位一次侧开关节点与输入滤波器，不是实验截图。公式中的dv/dt指寄生电容两端电压的变化率。20皮法和10伏每纳秒是教学假设，乘积为0.2安瞬时位移电流；不能把它当作LISN测得的噪声电流。最后强调电流必须有返回路径，为下一页共模回路铺垫。',R.r1+'\n'+R.r9+'，PDF第11页Fig.3a；原文[0039]说明图为概念示意。\n'+R.r10+'\n本组假设算例：20×10^-12 F ×10×10^9 V/s =0.2 A。');
}
// 04 Model reading is connected to an actual measurement distinction.
{
 const s=await slide('共模回路与传导干扰测量','图：[10] 第19页；共模／差模分离原理见其第17页');
 heading(s,'TI 共模教学等效模型',63,150,650);
 await img(s,'ti_common_mode_loop.png',60,203,669,346,'TI共模等效回路：噪声源、寄生电容、LISN和输入滤波器');
 para(s,'图中 GND 是共同测量参考平面；真实电源中**一次侧、二次侧、机壳与 PE 的节点关系**仍需逐一核对。',65,565,655,69,25);
 heading(s,'沿电流返回路径解释滤波',760,150,453);
 para(s,'VSW 经 CPAR 耦合到参考平面，再经 LISN 的测量终接及两条电源线返回噪声源。共模分量在 L1、L2 上同向，差模分量在两线间反向；共模电感主要提高前者的串联阻抗。',760,199,453,163,26);
 heading(s,'先分离模式，再评价改善',760,376,453,28);
 text(s,'V_CM = (V₁ + V₂) / 2\nV_DM = (V₂ − V₁) / 2',760,425,453,89,27,C.blue,true);
 para(s,'和、差运算需同一参考并保留相位信息，不能直接加减两路 dBμV。若主导超限分量是差模，仅改变共模支路未必有效。',760,528,453,120,25);
 notes(s,'70 秒','先沿左图讲一次完整回路，再说明共模与差模的区别。两条线相同参考下的和、差可用来分离模式；不能把两张幅值频谱直接相加就声称完成分离，因为相位关系也重要。该模型的GND是共同测量参考平面，实际隔离电源的各返回点不能凭符号合并。接下来比较传统滤波器的高频阻抗。',R.r10+'，第17页和第19页。\n'+R.r1);
}
// 05 Two original evidence images and interpretation, following reference slide 9/10 density.
{
 const s=await slide('传统滤波器的高频阻抗约束','图：[2] Fig.2、Fig.5a；该文建模范围 9 kHz–200 MHz');
 heading(s,'传统 EMI 滤波器实物与拓扑',63,151,646,28);
 await img(s,'aces_filter.png',60,199,658,282,'ACES Fig2传统滤波器实物与电路');
 para(s,'共模电感提供串联阻抗，抑制电容改变高频返回路径。体积与绕组、磁芯及绝缘结构有关；**标称电感值不能单独代表实际滤波性能**。',64,498,650,122,26);
 heading(s,'共模电感的测量与模型对比',760,151,458,28);
 await img(s,'aces_cm_impedance.png',763,195,453,265,'ACES Fig5a共模阻抗测量与GA及DE模型对比，裁去英文图注与白边');
 para(s,'图中阻抗随频率变化，在高频区出现明显下降和局部低谷。绕组寄生电容、漏感及磁芯特性使器件偏离理想 L 模型；该文比较 GA 与分段 DE 参数拟合。',760,469,458,139,25);
 emphasis(s,'设计含义：比较工作频段中的阻抗与抑制效果，不能只比较 L 值。',622);
 notes(s,'65 秒','这篇南师大团队论文提供了传统基线的全文证据。左图说明滤波器组成，右图给出共模电感测量阻抗及模型。这里不从缩略图读取精确参数，而是观察高频阻抗下降和低谷，解释理想电感模型为什么不够。该文9千赫到200兆赫是模型研究范围，不能等同于某产品的合规频段。',R.r2+'，PDF第3页Fig.2与第5页Fig.5a（期刊第93、95页）。');
}
// 06 Connection, electromechanical mechanism, and the LISN criterion all on the slide.
{
 const s=await slide('压电谐振支路的接入与作用','图：[9] Fig.3b，专利原理示意；机制依据 [3]、[8]');
 heading(s,'图中元件 10：压电抑制支路',63,150,659,28);
 await img(s,'patent_piezo_filter.png',60,203,658,333,'EP3591819A1 Fig3b压电抑制元件概念接入位置');
 para(s,'该示意把元件放在一次、二次侧之间的高频支路。跨隔离边界接入时，应核对**绝缘、额定电压和器件使用资格**。',64,555,653,83,25);
 heading(s,'电机械耦合形成选频阻抗',760,150,456,28);
 para(s,'外加电压激励机械振动，机械响应再影响电流，使端口阻抗在谐振附近下降。离开相关模态时，器件可表现出电容性特征，因而它的作用具有明显频率选择性。',760,200,456,160,26);
 heading(s,'低阻抗与噪声改善的关系',760,366,456,28);
 para(s,'有效支路应使目标共模电流更多地在局部返回，从而减少进入测量端的分量。若接入位置改变了其他路径，或频率没有匹配，端口低阻抗也不能直接推出 LISN 噪声降低。',760,416,456,157,26);
 text(s,'评价对象：测量端的传导干扰变化',760,595,456,37,25,C.red,true);
 notes(s,'60 秒','这页先与第3页传统电路对照，指出元件10的位置。解释电压与机械振动之间的耦合，不说材料把所有EMI吸收掉。核心判断是测量端电流是否减小；低阻抗只是器件特征，必须结合位置。最后说明跨隔离接入需要器件资格，普通压电陶瓷片不能直接当作安全电容替换。',R.r9+'，PDF第11页Fig.3b，原文[0039]。\n'+R.r3+'\n'+R.r8);
}
// 07 Native model parameter table paired with original model and impedance figure.
{
 const s=await slide('BVD 模型与端口阻抗特征','图：[8] Fig.2、Fig.3；本文 C₀ / Lₘ / Cₘ / Rₘ 对应原图 Cp / L / C / R');
 heading(s,'单模等效电路',66,151,550,29);
 await img(s,'bvd_original.png',76,202,190,195,'BVD等效模型，静态电容并联串联运动支路');
 table(s,[['参数','对应物理意义'],['C₀','静态端口电容'],['Lₘ、Cₘ','等效惯性与弹性'],['Rₘ','等效损耗']],[112,251],282,204,363,183,23);
 para(s,'静态电容与串联运动支路并联，端口电流由两部分构成。**RLC 是机械响应的等效参数**，并不表示实物中额外串接三只分立元件。',66,419,582,113,26);
 para(s,'单模模型主要描述某一振动模态附近；多模特征需要增加运动支路，并通过建模或测量提取参数。',66,548,582,82,25);
 heading(s,'谐振区间内的幅值与相位',703,151,515,29);
 await img(s,'bvd_impedance_original.png',711,204,493,293,'Boles Fig3谐振与反谐振之间的阻抗幅值及相位原理示意');
 para(s,'串联谐振附近出现阻抗低谷，反谐振附近出现高阻抗峰；两者之间可呈感性，外侧可呈容性。**滤波设计应同时看幅值、相位和接入回路**。',704,523,512,132,26);
 notes(s,'85 秒','先用左图建立参数含义，再用右图解释它为何有低谷和高峰。原图Cp在本次记为C0，运动支路参数加下标m。原理图只给定性趋势，下一页用一套明确假设参数计算。单模适用范围应与模态对应，不能用这组参数覆盖全部频带。',R.r8+'，公开稿PDF第2页Fig.2、Fig.3。');
}
const C0=1e-9,Cm=50e-12,Rm=20,fsHz=300e3,Lm=1/((2*Math.PI*fsHz)**2*Cm);
const fpHz=fsHz*Math.sqrt(1+Cm/C0);
function impedance(f){const w=2*Math.PI*f;const im=w*Lm-1/(w*Cm);const d=Rm*Rm+im*im;const yr=Rm/d,yi=w*C0-im/d;return 1/Math.hypot(yr,yi);}
const round6=n=>Number(n.toFixed(6));
// 08 Native calculated chart plus interpretation and explicit assumptions.
{
 const s=await slide('单模阻抗的谐振与反谐振','本组理想模型计算，非实测或论文拟合；模型依据 [8]');
 const freq=[...Array.from({length:145},(_,i)=>286+i*.25),round6(fpHz/1000)].sort((a,b)=>a-b);
 chartFont(s.charts.add('scatter',{position:{left:52,top:157,width:708,height:349},categories:freq.map(String),series:[{name:'单模 BVD',xValues:freq,values:freq.map(f=>round6(20*Math.log10(impedance(f*1000)))),line:{fill:C.accent,width:3,style:'solid'},marker:{symbol:'none'}}],scatterOptions:{style:'line'},hasLegend:false,xAxis:{title:'频率 / kHz',min:286,max:322,majorUnit:6,textStyle:{fontSize:19},majorGridlines:null},yAxis:{title:'20 log₁₀(|Z| / 1 Ω)',min:20,max:110,majorUnit:20,textStyle:{fontSize:19},majorGridlines:{fill:'#DDE3EA',width:1}},chartFill:'none',plotAreaFill:'none'}));
 heading(s,'计算参数与口径',67,526,665,27);
 text(s,'C₀ = 1 nF，Cₘ = 50 pF，Rₘ = 20 Ω\nfₛ = 300 kHz，Lₘ ≈ 5.63 mH',67,571,665,73,25);
 heading(s,'串联谐振：运动支路电抗抵消',796,155,421,27);
 para(s,'在 fₛ 附近，ωLₘ 与 1/(ωCₘ) 互相抵消，运动支路主要受 Rₘ 限制。该算例在 300 kHz 的端口阻抗约为 **20 Ω**，形成选频低阻抗窗口。',796,202,421,135,25);
 heading(s,'反谐振：导纳虚部近似相消',796,350,421,27);
 para(s,'低损耗近似给出 fₚ ≈ 307.4 kHz；在该频点算得阻抗约 **13.4 kΩ**。低谷附近同时存在高峰，接入后应检查目标频点及相邻频段。',796,398,421,128,25);
 para(s,'纵轴是以 1 Ω 为参考的阻抗幅值表达；它不是 dBμV，也不能直接读成系统 EMI 衰减量。',796,548,421,96,25);
 notes(s,'60 秒','这张原生数据图可以编辑。教学参数没有对应具体陶瓷器件。在300千赫，运动支路电抗抵消，端口阻抗约20欧；低损耗近似反谐振位置307.4千赫处约13.4千欧。这里是阻抗幅值的分贝表达，不能当作EMI电压衰减。重点说明低谷旁还有高峰，需要把接入后的完整系统重新检查。',R.r8+'\n本组公式：Z=[jωC0+1/(Rm+jωLm+1/(jωCm))]^-1；fs=300kHz,C0=1nF,Cm=50pF,Rm=20Ω,Lm=5.628954647mH。|Z(fs)|=19.985803 Ω；在近似fp=307.408523kHz处|Z|≈13.412kΩ，未把近似频率称为有损模型精确峰值。');
}
// 09 A calculation table connects the curves to the engineering tradeoff.
{
 const s=await slide('谐振匹配与频率偏差','运动支路理想计算示例，非 EMI 衰减实测；频移讨论见 [3]');
 const det=Array.from({length:81},(_,i)=>round6(-1+i*.025));
 chartFont(s.charts.add('scatter',{position:{left:52,top:155,width:708,height:349},categories:det.map(String),series:[50,200].map((q,i)=>({name:`Q = ${q}`,xValues:det,values:det.map(d=>{const r=1+d/100;return round6(Math.sqrt(1+q*q*(r-1/r)**2))}),line:{fill:i?C.red:C.accent,width:3,style:'solid'},marker:{symbol:'none'}})),scatterOptions:{style:'line'},hasLegend:true,legend:{position:'bottom',textStyle:{fontSize:21}},xAxis:{title:'频率偏差 / %',min:-1,max:1,majorUnit:.5,textStyle:{fontSize:19},majorGridlines:null},yAxis:{title:'|Zₘ| / Rₘ',min:0,max:5,majorUnit:1,crossesAt:-1,textStyle:{fontSize:19},majorGridlines:{fill:'#DDE3EA',width:1}},chartFill:'none',plotAreaFill:'none'}));
 table(s,[['相对频偏','Q = 50','Q = 200'],['0%','1.00','1.00'],['+1%','约 1.41','约 4.10']],[220,224,224],65,522,668,109,23);
 heading(s,'Q 值与频偏敏感性',796,155,421,28);
 para(s,'在谷值处 |Zₘ| = Rₘ；偏离谐振时，运动支路电抗的绝对值增加。较高 Q 表示更尖锐的谐振特征，因此**相同频偏下，阻抗相对谷值上升更快**。两条曲线仅比较该相对变化。',796,204,421,165,25);
 heading(s,'工程设计中的检查项目',796,383,421,28);
 para(s,'开关频率变化、器件温度和机械安装条件都应纳入频率匹配。多个模态可以匹配多个离散峰，但有效带宽和峰间覆盖仍需验证。',796,432,421,127,25);
 para(s,'完整端口还包含 C₀，外部还有回路阻抗；图中比例不能直接换算为传导干扰改善量。',796,565,421,88,24);
 notes(s,'80 秒','这里人为选择Q50和200，不沿用上一页的参数。对比+1%频偏，归一支路阻抗分别约1.41和4.10倍，说明高Q对匹配更敏感。由于纵轴只看运动支路，不是完整BVD或系统衰减，不能进一步声称噪声增加了多少分贝。应用时需要检查器件与噪声峰的漂移，多个模态也不自动等于宽频抑制。',R.r3+'\n'+R.r5+'\n本组公式：|Zm|/Rm=sqrt(1+Q²(r−1/r)²)，r=f/fs；Q为教学选择。+1%时Q50=1.410717,Q200=4.103898。');
}
function caseChart(s,categories,values,format='0%'){
 chartFont(s.charts.add('bar',{position:{left:51,top:222,width:713,height:327},categories,series:[{name:'该论文基线',values:categories.map(()=>1),fill:'#A7B4C9'},{name:'采用压电方案',values,fill:C.accent}],barOptions:{direction:'column',grouping:'clustered',gapWidth:95},hasLegend:true,legend:{position:'bottom',textStyle:{fontSize:20}},xAxis:{textStyle:{fontSize:21},majorGridlines:null},yAxis:{title:'相对该论文基线',min:0,max:1.1,majorUnit:.25,numberFormatCode:'0%',textStyle:{fontSize:19},majorGridlines:{fill:'#DDE3EA',width:1}},dataLabels:{showValue:true,position:'outEnd',numberFormatCode:format,textStyle:{fontSize:21}},chartFill:'none',plotAreaFill:'none'}));
}
// 10 Literal abstract numbers with scope, arithmetic and practical interpretation.
{
 const s=await slide('案例 A：滤波器体积与直流损耗','[4] APEC 2024，作者原始摘要；具体功率、器件配置与完整测量条件尚待全文核对');
 heading(s,'Hubert 等：反激变换器应用',65,152,672,29);
 caseChart(s,['滤波器体积','滤波器 DC 损耗'],[.36,.46]);
 para(s,'各指标以论文自身的传统方案为 100%：体积减少 64% 对应**剩余 36%**；滤波器 DC 损耗减少 54% 对应**剩余 46%**。两项指标的单位与分母分别计算。',65,559,668,92,24);
 heading(s,'论文摘要报告的结果',799,152,421,28);
 text(s,'体积减少 64%\nDC 损耗减少 54%',799,204,421,110,33,C.red,true);
 heading(s,'小型化的评价意义',799,334,421,28);
 para(s,'体积之外同时报告损耗，有助于讨论滤波器缩减是否伴随损耗代价。实际设计还应核对噪声抑制目标、比较基线和器件使用条件，才能判断结果是否适用于另一电源。',799,383,421,155,25);
 para(s,'统计对象为**滤波器**。整机还含功率级、散热和外壳；滤波器 DC 损耗减少不等于整机效率提升。',799,552,421,102,25);
 notes(s,'65 秒','先读两项结论，然后解释柱图剩余36%和46%的含义。这里两个指标分别归一化，不能相加。这项案例同时谈体积和损耗，但目前依据原始摘要，不把未取得的实验工况补出来。强调DC损耗属于滤波器，整机效率不能直接按54%提升。',R.r4+'\n算术重表达：1−0.64=0.36；1−0.54=0.46。基线属于本篇论文。');
}
// 11 Separate study, separate baseline, with multimode context.
{
 const s=await slide('案例 B：无源滤波器体积与重量','[5] JESTCH 2025，出版者原始摘要；几何尺寸、模型与多谐振频率为其研究内容');
 heading(s,'Zhang、Yan、Zhou：反激变换器应用',65,152,686,28);
 caseChart(s,['无源滤波器体积','无源滤波器重量'],[.0929,.1128],'0.00%');
 para(s,'对论文自身无源基线归一化：体积减少 90.71% 对应**剩余 9.29%**；重量减少 88.72% 对应**剩余 11.28%**。该图只重表达摘要数值，不表示与案例 A 统一对照。',65,559,668,92,24);
 heading(s,'论文摘要报告的结果',799,152,421,28);
 text(s,'体积减少 90.71%\n重量减少 88.72%',799,204,421,110,32,C.red,true);
 heading(s,'模型与几何设计的联系',799,334,421,28);
 para(s,'该研究关注压电器件建模，以及几何尺寸与多个谐振频率的关系。多模态提供多个选频特征，有助于讨论对不同干扰峰的匹配；具体覆盖频段仍需依据器件和系统测量。',799,383,421,155,25);
 para(s,'数据属于该案例的**无源滤波器**。基线组成和完整工况尚待全文核对，不能只按缩减比例判断方案优劣。',799,552,421,102,25);
 notes(s,'65 秒','该篇作者与授课老师研究方向有关，但评价仍看证据。出版者摘要报告体积与重量缩减，图上是各自剩余比例。多模设计是原摘要提到的研究内容，不代表自动覆盖整个传导频段。讲到这里自然转到下一页，说明两篇百分比不能直接排序。',R.r5+'\n算术重表达：1−0.9071=0.0929；1−0.8872=0.1128。原始出版者页面与DOAJ元数据核对书目及摘要。');
}
// 12 Dense native evidence comparison and two interpretation paragraphs.
{
 const s=await slide('文献结果的比较口径','依据 [4]、[5] 原始摘要整理；“待全文核对”表示目前没有足够的完整条件证据');
 table(s,[['比较维度','案例 A：APEC 2024','案例 B：JESTCH 2025'],['应用对象','反激变换器中的 EMI 滤波器','反激变换器中的无源滤波器'],['摘要指标','体积减少 64%\n滤波器 DC 损耗减少 54%','体积减少 90.71%\n重量减少 88.72%'],['比较基线','该论文自身的传统方案\n具体组成待全文核对','该论文自身的无源方案\n具体组成待全文核对'],['未确认条件','功率、频率、负载与器件参数\n完整传导干扰检测设置','功率、频率、负载与器件参数\n完整传导干扰检测设置'],['可支持的结论','该案例报告滤波器小型化\n并改善滤波器 DC 损耗','该案例报告无源滤波器\n体积与重量明显缩减']],[223,469,468],60,157,1160,348,23);
 heading(s,'本组比较原则',66,547,1150,28);
 para(s,'跨论文比较应统一电源功率与工况、开关频率、滤波器组成及体积计算方式，并核对测量频段、检波器、终接和评价目标。**减少比例更大，只能说明相对自身基线变化更大**。',66,589,1150,67,25);
 notes(s,'65 秒','两篇都能支持特定案例的小型化潜力，但目前不能支持同等条件下方法优劣排序。表中明确写出已确认和未确认字段。跨文献百分比的分母是关键；也不能把同作者组的后续文章当作完全独立验证。汇报只点出统计对象、基线和工况三项，详细字段留给观众阅读。',R.r4+'\n'+R.r5);
}
// 13 Structured engineering decision conditions with substantive actions.
{
 const s=await slide('压电方案的工程适用条件','本组基于 [3]、[5]、[7]、[8] 的机制与案例分析；表内为设计判断，不是额外实测结果');
 table(s,[['检查条件','对抑制机理的影响','采用前应核对的内容'],['噪声峰与谐振匹配','稳定离散峰适合选频抑制；\n频移可能使低阻抗窗口错开','开关频率变化、温度频移、\n目标峰频率及允许偏差'],['共模返回路径','有效支路应减少进入测量端\n的电流，位置决定分流效果','一次／二次侧与参考平面；\nLISN 端改善及相邻频段'],['宽频与多模需求','多个模态可针对多个峰，\n峰间覆盖不能直接推定','其他谐波、宽频噪声；\n是否仍需传统滤波配合'],['隔离与器件使用','跨隔离或接入 PE 时，\n功能效果之外还有使用约束','绝缘、额定电压、漏电流\n及相应器件资格'],['安装与稳定性','机械边界影响振动响应；\n参数变化影响谐振匹配','固定方式、温度范围、\n损耗及长期参数稳定性']],[231,458,471],60,157,1160,421,23);
 emphasis(s,'适用判断应同时覆盖：频率匹配、正确路径和实际使用条件。',604);
 notes(s,'75 秒','把本组分析归纳为三类。首先匹配噪声频率，其次接入正确的回路，最后检查隔离与机械条件。表里没有给出未测得的可靠性数字，而是指出应用前要验证的对象。如果噪声频谱分散或漂移，压电谐振可能需要与常规滤波配合。',R.r3+'\n'+R.r5+'\n'+R.r7+'\n'+R.r8);
}
// 14 Four substantive academic paragraphs replace three sparse slogan lines.
{
 const s=await slide('调研结论与进一步研究','依据 [2]–[5]、[8]、[9]；本次工作为文献调研与简化模型分析');
 heading(s,'1. 机理：以选频阻抗改变共模返回路径',66,152,1148,29);
 para(s,'压电器件的电机械谐振使端口在特定频率出现低阻抗。有效抑制依赖该特征与噪声峰匹配，并将电流引入合适的局部返回路径；单独观察器件阻抗不能替代系统传导干扰评价。',66,197,1148,69,26);
 rule(s,66,280,1148);
 heading(s,'2. 证据：特定反激案例报告明显的滤波器缩减',66,299,1148,29);
 para(s,'APEC 2024 摘要报告体积减少 64%、滤波器 DC 损耗减少 54%；JESTCH 2025 摘要报告无源滤波器体积减少 90.71%、重量减少 88.72%。两者各自使用不同基线，结论保留在对应统计对象与案例中。',66,344,1148,100,26);
 rule(s,66,455,1148);
 heading(s,'3. 条件：高 Q 带来频率匹配要求，应用还受回路约束',66,471,1148,29);
 para(s,'简化模型显示高 Q 支路对频偏更敏感；工程应用还应核对温度、安装、隔离与有效带宽。多模态可以增加选频特征，但仍需确认峰间覆盖及其他频段的干扰表现。',66,516,1148,69,26);
 text(s,'进一步研究：取得案例全文，补全工况与基线，并比较系统频谱及器件代价。',66,608,1148,48,26,C.red,true);
 notes(s,'30 秒','结束页只讲三句：压电谐振提供选频路径；文献案例报告滤波器缩减；推广需要频率、回路和使用条件一起验证。进一步工作是取得全文并补齐比较条件。谢谢老师和同学，后面有公式和问答备用页。',R.r2+'\n'+R.r3+'\n'+R.r4+'\n'+R.r5+'\n'+R.r8+'\n'+R.r9);
}
// 15 Bibliography with access status and the evidence each source contributes.
{
 const s=await slide('参考文献：核心研究与证据','编号与正文一致；完整英文题名、网址与引用定位均写在本页备注',true);
 table(s,[['编号','文献与出版信息','本次用途及证据状态'],['[1]','Xu et al. 传导 EMI 抑制方法综述\nIET Power Electronics, 2022\nDOI: 10.1049/pel2.12348','源头与传播路径的研究背景\n本次依据原始摘要'],['[2]','Pei et al. 共模电感高频建模与 EMI 抑制\nACES, 2023, 38(2):91–99\nDOI: 10.13052/2023.ACES.J.380203','传统滤波基线、阻抗测量\n已获取全文及原始图'],['[3]','Hubert et al. Piezoelectric EMI Filter\nIEEE TPEL, 2021, 36(6):6624–6643\nDOI: 10.1109/TPEL.2020.3043091','压电 EMI 滤波机制及频移讨论\n本次依据作者原始摘要'],['[4]','Hubert et al. 滤波器体积与损耗缩减\nAPEC, 2024:1506–1511\nDOI: 10.1109/APEC48139.2024.10509072','体积 −64%，滤波器 DC 损耗 −54%\n原始摘要；完整工况待核对'],['[5]','Zhang, Yan, Zhou. 压电滤波建模与性能评价\nJESTCH, 2025, 66:102060\nDOI: 10.1016/j.jestch.2025.102060','体积 −90.71%，重量 −88.72%\n出版者摘要；全文尚未取得']],[82,688,390],60,157,1160,478,22);
 notes(s,'备用','参考文献按本次实际用途区分证据。传统高频建模已读全文；核心压电案例采用原始作者或出版者摘要。没有把所有条目写成已读全文。详细网址可从备注或配套讲稿打开。',Object.values(R).slice(0,5).join('\n\n'));
}
{
 const s=await slide('参考文献：模型与扩展阅读','扩展文章与核心案例的作者组存在重合，应按研究内容评估其贡献',true);
 table(s,[['编号','文献与出版信息','本次用途及证据状态'],['[6]','Zhang, Yan, Zhou. 压电抑制元件建模与测量\nIEEE TIM, 2025, 74:1–11\nDOI: 10.1109/TIM.2025.3557100','后续建模与测量研究线索\n已核对书目；未据其补写数值'],['[7]','Zhang, Yan, Zhou. 压电陶瓷阻抗谐振设计\nAEU, 2025, 190:155622\nDOI: 10.1016/j.aeue.2024.155622','选频设计与多模研究线索\n本次依据出版者摘要'],['[8]','Boles et al. 压电功率电子：进展与挑战\nIEEJ, 2023, 12(3):254–263\nDOI: 10.1541/ieejjia.22008217','BVD、阻抗相位、振动模态\n已获取作者公开全文'],['[9]','Hubert, Dürbaum, Rupitsch. EMI 抑制元件\nEP3591819A1，公开日 2020-01-08\nElectromagnetic interference suppression components','反激电路与压电接入位置\n已获取公开专利；图为概念示意'],['[10]','Texas Instruments. Introduction to EMI\nin power supply design，公开教学资料\n共模与差模分离／LISN 等效回路','共模路径与测量模式解释\n已获取公开资料，第17／19页']],[82,688,390],60,157,1160,478,22);
 notes(s,'备用','三篇同作者组论文不能自动算作三个独立团队的重复验证。Boles综述用于模型与模态，专利用于原理接入，TI教学资料用于共模和测量解释。各类材料用途不同，不能互相替代实验数据。',Object.values(R).slice(5).join('\n\n'));
}
// 17 Derivation steps beside numeric results.
{
 const s=await slide('备用：BVD 公式与计算步骤','理想单模模型；反谐振频率采用低损耗近似；参数为本组教学选择',true);
 heading(s,'并联导纳与谐振条件',66,152,658,29);
 text(s,'Z(ω) = [jωC₀ + 1 / Zₘ(ω)]⁻¹\nZₘ(ω) = Rₘ + jωLₘ + 1/(jωCₘ)',66,204,672,111,29,C.blue,true);
 para(s,'运动支路串联谐振满足 ωLₘ − 1/(ωCₘ) = 0；反谐振则对应两支路导纳虚部近似相消。',66,330,672,82,26);
 text(s,'fₛ = 1 / (2π√(LₘCₘ))\nfₚ ≈ fₛ√(1 + Cₘ/C₀)\nQ = ωₛLₘ / Rₘ',66,434,672,148,29);
 heading(s,'第8页算例的复算结果',793,152,425,29);
 table(s,[['计算项','结果'],['Lₘ','约 5.629 mH'],['fₚ（近似）','约 307.409 kHz'],['|Z(fₛ)|','约 19.986 Ω'],['|Z(fₚ近似)|','约 13.412 kΩ'],['该组参数 Q','约 530.5']],[217,208],793,207,425,251,23);
 para(s,'有损端口的精确极值可能偏移；第9页另选 Q = 50、200 比较频偏，不沿用本组参数。单模适用范围应与对应振动模态一致。',793,480,425,125,25);
 emphasis(s,'模型参数、测量阻抗和系统衰减是三个不同层面的评价对象。',620);
 notes(s,'备用','如果问到推导，从并联导纳开始，先求串联支路电抗抵消，再求低损耗时总导纳虚部相消。Q530.5属于第8页参数，第9页Q50和200是单独的教学选择。表中307.409千赫为近似反谐振频率，并非有损端口精确峰值频率。',R.r8+'\n复算：Lm=1/((2π×300000)^2×50pF)=5.628954647mH；Q=2πfsLm/Rm=530.51648。');
}
// 18 A native scenario table strengthens the previously sparse percentage explanation.
{
 const s=await slide('备用：滤波器缩减与整机收益','本组简化口径分析；α 的数值为教学假设，64% 取案例 A 作算例',true);
 heading(s,'体积统计边界与简化推导',66,152,1148,29);
 para(s,'整机体积还包含功率级、控制、散热和外壳。设滤波器原占整机比例为 α，其他部分体积保持不变且各部分可加总，则：',66,201,1148,65,26);
 text(s,'r整机 = α · r滤波器      r 表示相对于各自原体积的缩减比例',66,284,1148,63,32,C.red,true);
 table(s,[['假设滤波器占比 α','滤波器缩减比例','估算整机缩减'],['10%','64%','6.4%'],['20%','64%','12.8%'],['30%','64%','19.2%']],[390,385,385],66,365,1148,172,25);
 para(s,'同理，滤波器损耗减少对整机损耗的影响，还取决于滤波器原损耗占比。实际外形受布局、散热与绝缘间距约束，器件体积之和也不必等于外壳体积。',66,560,1148,69,26);
 notes(s,'备用','用这页防止把滤波器缩减误说成整机同幅度缩减。20%占比乘64%等于12.8%，但20%是假设，不是案例A实测。另两行是同一简化模型的情景分析。损耗的分母也要区分滤波器与整机；若讨论效率，还要知道输入输出功率，不能直接把百分比当效率增幅。',R.r4+'\n本组推导：Vnew=Vother+(1−rfilter)Vfilter，alpha=Vfilter/Vtotal，故rtotal=alpha*rfilter。');
}
// 19 Answers contain reasons and direct links to the relevant part of the argument.
{
 const s=await slide('备用：常见问题与回答依据','答疑时先给短结论，再定位机制、测量条件或统计口径',true);
 table(s,[['常见问题','回答要点与依据'],['能取代全部传统滤波吗？','需按噪声峰、频偏和有效带宽判断；对宽频干扰可能仍需\n常规滤波配合。见第9、13页，不能由单一谐振低谷推出。'],['低阻抗为何未必降低 EMI？','低阻抗只有放在合适返回路径才可能减小测量端电流；\n应同时看源、支路与LISN。见第4、6页。'],['单峰衰减能证明合规吗？','需要相应频段、限值、负载工况及检波方式的完整测量；\n某一频点改善不能代替全频段评价。'],['dB 与 dBμV 如何解释？','dB用于表示比值或改善量；dBμV为相对1 μV的电压电平。\n第8页阻抗纵轴以1 Ω为参考，应按其定义解释。'],['缩减数字为何不能直接排序？','两篇采用各自滤波器基线，统计量及工况未统一；\n滤波器体积也不同于整机体积。见第12、18页。'],['主动 EMI 滤波怎样比较？','还需考察辅助供电、环路及带宽；本次主线是压电谐振，\n未开展统一条件下的跨方法性能评估。']],[368,792],60,157,1160,478,23);
 notes(s,'备用','答疑不要把限制简单回答为不能。先解释为什么，再返回对应页。接入路径、频率带宽、合规测量和百分比分母分别属于不同问题。主动滤波可以作为相关路线简述，但不补写没有开展的性能对比。',R.r1+'\n'+R.r3+'\n'+R.r10);
}
// 20 Original mode illustration plus physical interpretation and modeling consequence.
{
 const s=await slide('备用：振动模态与多模设计','图：[8] Fig.12；几何尺寸与多谐振设计讨论见 [5]、[7]',true);
 para(s,'不同几何形状、极化方向与机械边界，可形成不同的振动模态。图中展示四类常见模态；具体谐振频率仍需结合材料参数与器件尺寸确定。',65,153,1150,65,26);
 await img(s,'vibration_modes_original.png',65,248,1150,183,'Boles Fig12厚度伸缩、厚度剪切、轮廓伸缩和径向振动');
 const labels=['厚度伸缩','厚度剪切','轮廓伸缩','径向振动'];labels.forEach((v,i)=>text(s,v,65+288*i,446,280,41,27,C.accent,true,'center'));
 heading(s,'多模模型的含义',65,508,558,28);
 para(s,'可并联多个运动支路，分别描述不同模态的谐振与损耗。参数应来自器件建模或测量，而非任意增加谐振点。',65,553,558,96,25);
 heading(s,'对滤波设计的启示',660,508,555,28);
 para(s,'多个模态可以针对多个离散干扰峰。**多峰不等于连续宽带抑制**；峰间覆盖、频移与接入后的系统频谱仍需逐项验证。',660,553,555,96,25);
 notes(s,'备用','四个模态用于说明几何与极化如何影响可激发的机械响应，不用这张示意图读出具体器件频率。多模BVD可以扩展运动支路，但每个支路参数都需要依据。对滤波而言，多个峰只能说明提供多个选频特征，不能自动称宽带方案。',R.r8+'，作者公开稿第5页Fig.12。\n'+R.r5+'\n'+R.r7);
}

await fs.mkdir(buildDir,{recursive:true});
await fs.mkdir(`${buildDir}/draft_previews`,{recursive:true});
await fs.mkdir(`${workspaceDir}/输出`,{recursive:true});
const candidatePath=`${buildDir}/candidate.pptx`;
await (await PresentationFile.exportPptx(p)).save(candidatePath);
await fs.writeFile(`${buildDir}/notes.json`,JSON.stringify(slides.map((s,i)=>({slide:i+1,notes:s.speakerNotes.text})),null,2));
const nativeTableOwners=[2,7,9,12,13,15,16,17,18,19];
const result=await finalizePresentation({workspaceDir,candidatePath,finalPath,pythonExecutable:RUNTIME_PYTHON,integrityValidatorPath:`${SKILL_DIR}/container_tools/inspect_presentation_package_integrity.py`,layoutValidatorPath:`${SKILL_DIR}/container_tools/inspect_presentation_layout_geometry.py`,layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...nativeTableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],requiredNativeTableOwnerSlides:nativeTableOwners,requiredNativeChartOwnerSlides:[8,9,10,11],fontPolicy,materializeLiteralChartWorkbooks:true,verifyArtifactToolImport:true,receiptPath:`${buildDir}/${finalName}.v2.validation.json`});
console.log(JSON.stringify({finalPath:result.finalPath,slideCount:result.packageIntegrity.slide_count,bytes:result.byteCount,sha256:result.finalSha256,package:result.packageIntegrity.status,layoutFindings:result.presentationLayout.finding_count,chartValidation:result.nativeChartValidation.passed}));
