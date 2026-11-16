const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/5a4c245b-5f22-49bb-9eac-e803fbbf62d2_cc444313-77b5-4ddb-8a26-34b769da2020/pptx/scripts/apply_theme.js");

const THEME = { name: "Gopal", headFontFace: "Yu Gothic", bodyFontFace: "Yu Gothic",
  colors: { dk1:"222222", lt1:"FFFFFF", dk2:"3A3A4A", lt2:"F6F0FA",
    accent1:"D6409F", accent2:"7B5CF0", accent3:"1FA8D6", accent4:"F4B860", accent5:"666677", accent6:"C0392B",
    hlink:"D6409F", folHlink:"666677" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "福利厚生の案(健康・防災)";
pres.company = "株式会社Gopal";
const C = pres.SchemeColor;
const HEALTH = C.accent3, DISASTER = C.accent1;

async function icon(name, color="FFFFFF") {
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: "#"+color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256,256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const footer = [
  { text:{ text:"Gopal Confidential", options:{ x:0.35, y:5.2, w:3, h:0.3, fontSize:11, color:C.text1, margin:0 } } },
  { text:{ text:"©Gopal Co., Ltd. 2026. All rights reserved.", options:{ x:5.2, y:5.2, w:3.9, h:0.3, fontSize:11, color:C.text1, align:"right", margin:0 } } },
];
pres.defineSlideMaster({ title:"CONTENT", background:{ path:"content_bg.png" },
  objects:[
    { rect:{ x:0.3, y:0.2, w:0.09, h:0.5, fill:{ color:C.accent5 } } },
    { image:{ path:"gopal_logo.png", x:8.2, y:0.1, w:1.4, h:0.56 } },
    { placeholder:{ options:{ name:"title", type:"title", x:0.55, y:0.15, w:7.5, h:0.6, fontSize:26, bold:true, color:C.text1, valign:"middle", align:"left", margin:0 }, text:"" } },
    ...footer ],
  slideNumber:{ x:9.15, y:5.2, w:0.5, h:0.3, fontSize:11, color:C.text1 } });
pres.defineSlideMaster({ title:"COVER", background:{ path:"cover_bg.png" },
  objects:[ { image:{ path:"gopal_logo.png", x:7.6, y:0.4, w:2.0, h:0.8 } }, ...footer ] });

const slideTitle = (s,t)=> s.addText(t,{placeholder:"title"});
const tb = (s,t,o)=> s.addText(t,Object.assign({isTextBox:true,margin:0},o));

(async()=>{
// 1 cover
pres.addSection({title:"導入"});
let s = pres.addSlide({masterName:"COVER", sectionTitle:"導入"});
s.addShape(pres.ShapeType.ellipse,{x:0.7,y:1.35,w:0.7,h:0.7,fill:{color:HEALTH},line:{type:"none"},objectName:"HealthIconBg"});
s.addImage({data:await icon("FaHeartbeat"),x:0.88,y:1.53,w:0.34,h:0.34,altText:"健康"});
s.addShape(pres.ShapeType.ellipse,{x:1.55,y:1.35,w:0.7,h:0.7,fill:{color:DISASTER},line:{type:"none"},objectName:"DisasterIconBg"});
s.addImage({data:await icon("FaShieldAlt"),x:1.73,y:1.53,w:0.34,h:0.34,altText:"防災"});
tb(s,"福利厚生の案",{x:0.7,y:2.3,w:8,h:0.9,fontSize:44,bold:true,color:C.text1,objectName:"Title"});
tb(s,"健康・防災に関する福利厚生",{x:0.7,y:3.2,w:8,h:0.6,fontSize:26,color:C.accent1,objectName:"Subtitle"});
tb(s,"株式会社Gopal  オペレーション",{x:0.7,y:4.3,w:6,h:0.35,fontSize:14,color:C.accent5,objectName:"Org"});
s.addNotes("福利厚生として検討したい「健康」「防災」の2分野について、案を整理した資料です。");

const HEALTH_ITEMS=[
 ["FaUserMd","健康診断・人間ドック補助","病気を早く見つけるための受診を、会社が支える",
  "年1回の健康診断や人間ドックの受診費用を、会社が補助する制度です。",
  "忙しくて受診を後回しにしがちな人も、病気を早期に見つけやすくなります。",
  "提携医療機関での受診、または自己負担分を上限額の範囲内で精算します。"],
 ["FaBrain","メンタルヘルス相談窓口","話しにくい悩みを、外部の専門家に相談できる",
  "外部のカウンセラーなどに、匿名・無料で相談できる窓口を用意します。",
  "仕事のストレスや人間関係、家庭の悩みなど、社内には話しにくいことを相談できます。",
  "電話・チャット・オンラインで相談。相談内容が会社に伝わらない仕組みにします。"],
 ["FaSyringe","予防接種の費用補助","流行期の欠勤や配送の遅れを減らす",
  "インフルエンザなどの予防接種にかかる費用を補助します。",
  "流行期の体調不良による欠勤や、配送の遅れを減らせます。",
  "接種後に領収書を提出し、上限額の範囲内で補助します。"],
 ["FaRunning","運動・フィットネス補助","運転中心の生活で不足しがちな運動を後押し",
  "ジム、スポーツクラブ、ヨガ教室などの利用料を一部補助します。",
  "運転が中心で運動不足になりがちな人の体づくりや、腰痛の予防に役立ちます。",
  "月額の上限を決めて補助、またはオンライン運動サービスを法人で契約します。"],
 ["FaSun","腰痛・熱中症対策グッズ","仕事中の体の負担を、道具で軽くする",
  "腰サポーター、塩タブレット、冷却タオル、ネッククーラーなどを配布します。",
  "荷物の積み下ろしや、夏場の車内・屋外での作業で体の負担を減らせます。",
  "入社時や季節の変わり目に配布。希望者には追加で渡します。"],
 ["FaBed","健康セミナー・睡眠改善","体調管理の知識を、好きな時間に学べる",
  "体調管理、食事、睡眠、ストレッチなどの講座をオンラインで提供します。",
  "勤務時間が不規則になりやすく、睡眠不足や生活習慣の乱れが起きやすい人に向いています。",
  "短い動画やオンライン講座を、自分の好きな時間に視聴できます。"]];
const DISASTER_ITEMS=[
 ["FaMobileAlt","安否確認システム","災害時に、全員の無事をすばやく把握する",
  "災害時にスマホへ一斉通知し、「無事」「要支援」などを回答してもらう仕組みです。",
  "地震や豪雨で連絡が取りにくいときに、全員の状況をすぐに把握できます。",
  "アプリやメールで通知。年に数回の訓練配信で、使い方に慣れておきます。"],
 ["FaBoxOpen","防災グッズの支給","配送中も自宅でも、最低限の備えを持てる",
  "車載用(ライト、携帯トイレ、ブランケットなど)と自宅用の防災セットを配布します。",
  "配送中や自宅で被災したときに、当面を乗り切る備えになります。",
  "入社時に配布し、使用期限を管理して定期的に入れ替えます。"],
 ["FaUtensils","非常食・飲料水の備蓄","拠点で待機することになっても安心",
  "営業所や拠点に、水、非常食、簡易トイレなどを備蓄します。",
  "帰宅が難しくなったときや、拠点で待機が必要になったときに役立ちます。",
  "人数と日数を決めて備蓄し、期限を管理する担当者を決めます。"],
 ["FaChalkboardTeacher","防災研修・避難訓練","いざという時に、慌てず行動できるようにする",
  "災害時の行動、避難経路、配送中に被災したときの判断方法を学びます。",
  "運転中や配送先で地震・豪雨に遭ったときに、落ち着いて行動できます。",
  "年1回の研修と避難訓練に、動画教材を組み合わせます。"],
 ["FaHandHoldingHeart","災害見舞金・特別休暇","被災後の生活再建を、お金と時間で支える",
  "自宅が被害を受けたときなどに、見舞金の支給や特別休暇を認める制度です。",
  "被災後の片付けや生活の立て直しに専念できます。",
  "被害の程度に応じた金額と休暇日数を、社内規程で定めます。"],
 ["FaCar","悪天候時の運行ルール","無理な運行によるけがや事故を防ぐ",
  "大雨・大雪・地震・台風のときに、運行の中止・待機・再開を判断する基準と連絡の流れを整えます。",
  "無理な運行による事故やけがを防ぎ、安全を最優先にできます。",
  "警報のレベルごとの対応表を作り、全員に共有します。"]];

async function listSlide(title,col,items,notes){
  const sl = pres.addSlide({masterName:"CONTENT", sectionTitle:title.slice(0,2)});
  slideTitle(sl,title);
  for (let i=0;i<6;i++){
    const cx=0.5+(i%3)*3.1, cy=1.15+Math.floor(i/3)*2.0, w=2.8, h=1.85;
    sl.addShape(pres.ShapeType.roundRect,{x:cx,y:cy,w,h,rectRadius:0.1,fill:{color:C.background1,transparency:15},line:{color:"E4D9EE",width:0.75},objectName:"Card"+i});
    sl.addShape(pres.ShapeType.ellipse,{x:cx+0.2,y:cy+0.18,w:0.5,h:0.5,fill:{color:col},line:{type:"none"},objectName:"CardIconBg"+i});
    sl.addImage({data:await icon(items[i][0]),x:cx+0.32,y:cy+0.30,w:0.26,h:0.26,altText:items[i][1]});
    tb(sl,String(i+1),{x:cx+w-0.5,y:cy+0.15,w:0.3,h:0.5,fontSize:24,bold:true,color:col,align:"right",objectName:"CardNo"+i});
    tb(sl,items[i][1],{x:cx+0.2,y:cy+0.78,w:w-0.4,h:0.32,fontSize:14,bold:true,color:C.text1,objectName:"CardTitle"+i});
    tb(sl,items[i][2],{x:cx+0.2,y:cy+1.1,w:w-0.4,h:0.65,fontSize:11,color:C.text1,valign:"top",objectName:"CardDesc"+i});
  }
  sl.addNotes(notes);
}
const NUM="①②③④⑤⑥";
async function detailSlide(prefix,idx,col,it){
  const sl = pres.addSlide({masterName:"CONTENT", sectionTitle:prefix});
  slideTitle(sl,prefix+NUM[idx]+" "+it[1]);
  sl.addShape(pres.ShapeType.roundRect,{x:0.5,y:1.2,w:2.9,h:3.8,rectRadius:0.12,fill:{color:C.background1,transparency:15},line:{color:"E4D9EE",width:0.75},objectName:"SummaryPanel"});
  sl.addShape(pres.ShapeType.ellipse,{x:1.45,y:1.5,w:1.0,h:1.0,fill:{color:col},line:{type:"none"},objectName:"SummaryIconBg"});
  sl.addImage({data:await icon(it[0]),x:1.72,y:1.77,w:0.46,h:0.46,altText:it[1]});
  tb(sl,it[1],{x:0.7,y:2.7,w:2.5,h:0.8,fontSize:18,bold:true,color:C.text1,align:"center",valign:"middle",objectName:"SummaryName"});
  tb(sl,it[2],{x:0.7,y:3.55,w:2.5,h:1.2,fontSize:13,color:C.accent5,align:"center",valign:"top",objectName:"SummaryCatch"});
  const labels=["内容","役立つ場面","実施のイメージ"];
  for (let k=0;k<3;k++){
    const y=1.2+k*1.3;
    sl.addShape(pres.ShapeType.roundRect,{x:3.7,y,w:1.5,h:0.34,rectRadius:0.17,fill:{color:col},line:{type:"none"},objectName:"Label"+k});
    tb(sl,labels[k],{x:3.7,y,w:1.5,h:0.34,fontSize:12,bold:true,color:C.background1,align:"center",valign:"middle",objectName:"LabelText"+k});
    tb(sl,it[3+k],{x:3.7,y:y+0.43,w:5.8,h:0.75,fontSize:14,color:C.text1,valign:"top",objectName:"Body"+k});
  }
}

pres.addSection({title:"健康"});
await listSlide("健康に関する福利厚生",HEALTH,HEALTH_ITEMS.map(a=>[a[0],a[1],a[2]]),"健康に関する福利厚生の案を6つ紹介します。各案の内容は次のページ以降で説明します。");
for (let i=0;i<6;i++) await detailSlide("健康",i,HEALTH,HEALTH_ITEMS[i]);
pres.addSection({title:"防災"});
await listSlide("防災に関する福利厚生",DISASTER,DISASTER_ITEMS.map(a=>[a[0],a[1],a[2]]),"防災に関する福利厚生の案を6つ紹介します。各案の内容は次のページ以降で説明します。");
for (let i=0;i<6;i++) await detailSlide("防災",i,DISASTER,DISASTER_ITEMS[i]);

await pres.writeFile({fileName:"welfare_health_disaster.pptx"});
await applyTheme("welfare_health_disaster.pptx",THEME);
})();
