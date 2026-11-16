const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/5a4c245b-5f22-49bb-9eac-e803fbbf62d2_cc444313-77b5-4ddb-8a26-34b769da2020/pptx/scripts/apply_theme.js");

const THEME = { name: "Gopal Welfare", headFontFace: "Yu Gothic", bodyFontFace: "Yu Gothic",
  colors: { dk1:"1F2D2E", lt1:"FFFFFF", dk2:"0B5D5E", lt2:"EAF4F2", accent1:"0B5D5E", accent2:"E8801A",
    accent3:"2A9D8F", accent4:"F4B860", accent5:"5F7274", accent6:"C0392B", hlink:"0B5D5E", folHlink:"5F7274" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "福利厚生の案(健康・防災)";
const C = pres.SchemeColor;

async function icon(name, color="FFFFFF") {
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: "#"+color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256,256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

pres.defineSlideMaster({ title:"CONTENT", background:{color:"FFFFFF"},
  objects:[{ placeholder:{ options:{ name:"title", type:"title", x:0.5, y:0.3, w:9, h:0.8, fontSize:30, bold:true, color:C.accent1, valign:"middle", align:"left", margin:0 }, text:"" } }],
  slideNumber:{ x:9.1, y:5.25, w:0.5, h:0.25, fontSize:10, color:C.accent5 } });
pres.defineSlideMaster({ title:"DARK", background:{color:"0B5D5E"} });

const slideTitle = (s,t)=> s.addText(t,{placeholder:"title"});
const tb = (s,t,o)=> s.addText(t,Object.assign({isTextBox:true,margin:0},o));

(async()=>{
// 1 title
pres.addSection({title:"導入"});
let s = pres.addSlide({masterName:"DARK", sectionTitle:"導入"});
s.addShape(pres.ShapeType.ellipse,{x:6.9,y:-0.9,w:4.6,h:4.6,fill:{color:C.accent3,transparency:70},line:{type:"none"},objectName:"Deco1"});
s.addShape(pres.ShapeType.ellipse,{x:7.6,y:2.9,w:3.4,h:3.4,fill:{color:C.accent2,transparency:60},line:{type:"none"},objectName:"Deco2"});
s.addImage({data:await icon("FaHeartbeat"),x:0.7,y:1.1,w:0.6,h:0.6,altText:"健康"});
s.addImage({data:await icon("FaShieldAlt","F4B860"),x:1.45,y:1.1,w:0.6,h:0.6,altText:"防災"});
tb(s,"福利厚生のご提案",{x:0.7,y:1.9,w:7,h:0.9,fontSize:44,bold:true,color:C.background1,objectName:"Title"});
tb(s,"健康・防災の取り組み案",{x:0.7,y:2.85,w:7,h:0.6,fontSize:26,color:C.accent4,objectName:"Subtitle"});
tb(s,"株式会社Gopal  オペレーション",{x:0.7,y:4.6,w:6,h:0.35,fontSize:14,color:C.background2,objectName:"Org"});
s.addNotes("福利厚生として検討したい「健康」「防災」の2分野について、案を整理した資料です。");

// 2 背景
pres.addSection({title:"背景と案"});
s = pres.addSlide({masterName:"CONTENT", sectionTitle:"背景と案"});
slideTitle(s,"なぜ今、健康と防災なのか");
const why=[["FaTruck","体が資本の仕事","長時間の運転・荷物の積み下ろしで、腰痛や疲労、睡眠不足が起きやすい"],
 ["FaCloudSunRain","移動中に被災するリスク","配送中に地震・豪雨・猛暑に遭う可能性があり、会社の備えが安心につながる"],
 ["FaHandsHelping","安心して長く働ける環境","福利厚生の充実は、定着率の向上や採用での魅力づくりにもつながる"]];
for (let i=0;i<3;i++){ const y=1.4+i*1.25;
  s.addShape(pres.ShapeType.roundRect,{x:0.5,y,w:9,h:1.05,rectRadius:0.12,fill:{color:C.background2},line:{type:"none"},objectName:"WhyCard"+i});
  s.addShape(pres.ShapeType.ellipse,{x:0.75,y:y+0.17,w:0.7,h:0.7,fill:{color:C.accent1},line:{type:"none"},objectName:"WhyIconBg"+i});
  s.addImage({data:await icon(why[i][0]),x:0.92,y:y+0.34,w:0.36,h:0.36,altText:why[i][1]});
  tb(s,why[i][1],{x:1.75,y:y+0.12,w:7.5,h:0.35,fontSize:18,bold:true,color:C.text1,objectName:"WhyHead"+i});
  tb(s,why[i][2],{x:1.75,y:y+0.5,w:7.5,h:0.45,fontSize:14,color:C.text1,valign:"top",objectName:"WhyBody"+i});
}
s.addNotes("背景の整理。数値データは社内の実態に合わせて追記できます。");

// 3 overview
s = pres.addSlide({masterName:"CONTENT", sectionTitle:"背景と案"});
slideTitle(s,"案の全体像:各6案");
const H=["健康診断・人間ドック補助","メンタルヘルス相談窓口","予防接種の費用補助","運動・フィットネス補助","腰痛・熱中症対策グッズ","健康セミナー・睡眠改善"];
const D=["安否確認システム","防災グッズの支給","非常食・飲料水の備蓄","防災研修・避難訓練","災害見舞金・特別休暇","悪天候時の運行ルール整備"];
const panels=[["健康",H,0.5,C.accent1,"FaHeartbeat"],["防災",D,5.15,C.accent2,"FaShieldAlt"]];
for (const [nm,list,x,col,ic] of panels){
  s.addShape(pres.ShapeType.roundRect,{x,y:1.3,w:4.35,h:3.85,rectRadius:0.12,fill:{color:C.background2},line:{type:"none"},objectName:nm+"Panel"});
  s.addShape(pres.ShapeType.ellipse,{x:x+0.25,y:1.5,w:0.6,h:0.6,fill:{color:col},line:{type:"none"},objectName:nm+"IconBg"});
  s.addImage({data:await icon(ic),x:x+0.40,y:1.65,w:0.3,h:0.3,altText:nm});
  tb(s,nm+"に関する福利厚生",{x:x+1.0,y:1.5,w:3.2,h:0.6,fontSize:20,bold:true,color:col,valign:"middle",objectName:nm+"Head"});
  s.addText(list.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<list.length-1}})),
    {isTextBox:true,x:x+0.3,y:2.3,w:3.8,h:2.7,fontSize:15,color:C.text1,valign:"top",paraSpaceAfter:8,margin:0,objectName:nm+"List"});
}

// 4/5 idea grids
async function grid(title,col,items,sec){
  const sl = pres.addSlide({masterName:"CONTENT", sectionTitle:sec});
  slideTitle(sl,title);
  for (let i=0;i<6;i++){
    const cx=0.5+(i%3)*3.1, cy=1.3+Math.floor(i/3)*1.97, w=2.8, h=1.8;
    sl.addShape(pres.ShapeType.roundRect,{x:cx,y:cy,w,h,rectRadius:0.1,fill:{color:C.background2},line:{type:"none"},objectName:"Card"+i});
    sl.addShape(pres.ShapeType.ellipse,{x:cx+0.2,y:cy+0.18,w:0.5,h:0.5,fill:{color:col},line:{type:"none"},objectName:"CardIconBg"+i});
    sl.addImage({data:await icon(items[i][0]),x:cx+0.32,y:cy+0.30,w:0.26,h:0.26,altText:items[i][1]});
    tb(sl,String(i+1),{x:cx+w-0.5,y:cy+0.15,w:0.3,h:0.5,fontSize:24,bold:true,color:col,align:"right",objectName:"CardNo"+i});
    tb(sl,items[i][1],{x:cx+0.2,y:cy+0.78,w:w-0.4,h:0.32,fontSize:14,bold:true,color:C.text1,objectName:"CardTitle"+i});
    tb(sl,items[i][2],{x:cx+0.2,y:cy+1.1,w:w-0.4,h:0.62,fontSize:11,color:C.text1,valign:"top",objectName:"CardDesc"+i});
  }
  return sl;
}
pres.addSection({title:"健康"});
s = await grid("健康に関する案",C.accent1,[
 ["FaUserMd","健康診断・人間ドック補助","年1回の受診費用を補助し、病気の早期発見を後押しする"],
 ["FaBrain","メンタルヘルス相談窓口","外部の専門家に匿名・無料で相談できる窓口を設ける"],
 ["FaSyringe","予防接種の費用補助","インフルエンザ等の接種費用を補助し、欠勤を防ぐ"],
 ["FaRunning","運動・フィットネス補助","ジムやスポーツ施設の利用料を一部補助する"],
 ["FaSun","腰痛・熱中症対策グッズ","サポーター、塩タブレット、冷却グッズを配布する"],
 ["FaBed","健康セミナー・睡眠改善","体調管理や睡眠の講座をオンラインで提供する"]],"健康");
s.addNotes("ドライバー業務に直結する腰痛・熱中症・睡眠・メンタルに焦点を当てた案です。");
pres.addSection({title:"防災"});
s = await grid("防災に関する案",C.accent2,[
 ["FaMobileAlt","安否確認システム","災害時にスマホで安否を一斉確認し、状況を把握する"],
 ["FaBoxOpen","防災グッズの支給","車載用・自宅用の防災セットを配布する"],
 ["FaUtensils","非常食・飲料水の備蓄","営業所や拠点に水と非常食を備蓄する"],
 ["FaChalkboardTeacher","防災研修・避難訓練","配送中に被災した場合の行動を学び、訓練する"],
 ["FaHandHoldingHeart","災害見舞金・特別休暇","被災時に見舞金と特別休暇で生活を支える"],
 ["FaCar","悪天候時の運行ルール","大雨・地震時の運行判断基準と連絡フローを整える"]],"防災");
s.addNotes("配送中の被災を想定した備えと、被災後の生活支援の両面で整理しました。");

// 7 priority table
pres.addSection({title:"進め方"});
s = pres.addSlide({masterName:"CONTENT", sectionTitle:"進め方"});
slideTitle(s,"優先して検討したい6案");
const hd=(t)=>({text:t,options:{bold:true,color:"FFFFFF",fill:{color:"0B5D5E"},align:"center",valign:"middle"}});
const cell=(t,al="center")=>({text:t,options:{align:al,valign:"middle",color:"1F2D2E"}});
const rows=[[ "案","期待効果","費用感","導入しやすさ"],
 ["安否確認システム","高","低","高"],["メンタルヘルス相談窓口","高","低〜中","高"],["予防接種の費用補助","中","低","高"],
 ["防災グッズの支給","中","低〜中","高"],["健康診断・人間ドック補助","高","中","中"],["災害見舞金・特別休暇","高","中","中"]];
s.addTable(rows.map((r,i)=> i==0? r.map(hd): r.map((t,j)=>({text:t,options:{align:j?"center":"left",valign:"middle",color:"1F2D2E",bold:j==0,fill:{color:i%2?"FFFFFF":"EAF4F2"}}}))),
 {x:0.5,y:1.3,w:9,colW:[3.9,1.7,1.7,1.7],rowH:0.5,fontSize:14,border:{type:"solid",pt:0.5,color:"CFE3DF"},objectName:"PriorityTable"});
tb(s,"※評価は当社想定の目安です。見積もり取得後に見直します",{x:0.5,y:4.95,w:8.5,h:0.3,fontSize:11,color:C.accent5,objectName:"Note"});

// 8 roadmap
s = pres.addSlide({masterName:"CONTENT", sectionTitle:"進め方"});
slideTitle(s,"導入ロードマップ(案)");
const ph=[["第1段階","〜3か月",C.accent1,["安否確認システム","メンタル相談窓口","予防接種の費用補助"]],
 ["第2段階","3〜6か月",C.accent3,["防災グッズの支給","防災研修・訓練","健康セミナー"]],
 ["第3段階","6〜12か月",C.accent2,["健康診断・人間ドック補助","災害見舞金制度","フィットネス補助"]]];
ph.forEach(([a,b,col,list],i)=>{
  const x=0.5+i*3.05;
  s.addShape(pres.ShapeType.chevron,{x,y:1.35,w:3.0,h:0.8,fill:{color:col},line:{type:"none"},objectName:"Chevron"+i});
  tb(s,a+"  "+b,{x:x+0.45,y:1.35,w:2.2,h:0.8,fontSize:15,bold:true,color:C.background1,valign:"middle",align:"center",objectName:"PhaseLabel"+i});
  s.addShape(pres.ShapeType.roundRect,{x:x+0.05,y:2.4,w:2.8,h:1.9,rectRadius:0.1,fill:{color:C.background2},line:{type:"none"},objectName:"PhaseCard"+i});
  s.addText(list.map((t,k)=>({text:t,options:{bullet:true,breakLine:k<list.length-1}})),
    {isTextBox:true,x:x+0.3,y:2.65,w:2.35,h:1.5,fontSize:14,color:C.text1,valign:"top",paraSpaceAfter:10,margin:0,objectName:"PhaseList"+i});
});
s.addNotes("低コスト・導入しやすい案から先行し、費用のかかる制度は予算確認後に進めます。");

// 9 next steps
s = pres.addSlide({masterName:"DARK", sectionTitle:"進め方"});
tb(s,"ご相談したいこと",{x:0.5,y:0.4,w:9,h:0.8,fontSize:32,bold:true,color:C.background1,valign:"middle",objectName:"Title"});
const q=[["FaUsers","対象者の範囲","社員のみか、ドライバー(業務委託含む)まで広げるか"],
 ["FaYenSign","予算の枠","年間・一人あたりの上限をどこに置くか"],
 ["FaListOl","優先順位","6案のうち、まず始める案はどれか"],
 ["FaCalendarCheck","担当と開始時期","担当部署の決定と、第1段階の開始月"]];
for (let i=0;i<4;i++){ const cx=0.5+(i%2)*4.6, cy=1.5+Math.floor(i/2)*1.7;
  s.addShape(pres.ShapeType.roundRect,{x:cx,y:cy,w:4.4,h:1.5,rectRadius:0.12,fill:{color:C.background1,transparency:90},line:{color:C.background1,width:0.75,transparency:60},objectName:"QCard"+i});
  s.addShape(pres.ShapeType.ellipse,{x:cx+0.25,y:cy+0.25,w:0.6,h:0.6,fill:{color:C.accent2},line:{type:"none"},objectName:"QIconBg"+i});
  s.addImage({data:await icon(q[i][0]),x:cx+0.40,y:cy+0.40,w:0.3,h:0.3,altText:q[i][1]});
  tb(s,q[i][1],{x:cx+1.1,y:cy+0.2,w:3.1,h:0.4,fontSize:18,bold:true,color:C.accent4,objectName:"QHead"+i});
  tb(s,q[i][2],{x:cx+1.1,y:cy+0.65,w:3.1,h:0.75,fontSize:14,color:C.background1,valign:"top",objectName:"QBody"+i});
}
await pres.writeFile({fileName:"welfare_health_disaster.pptx"});
await applyTheme("welfare_health_disaster.pptx",THEME);
})();
