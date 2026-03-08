function normalizeAddressText(address){
  return String(address || "")
    .trim()
    .replace(/[　\s]+/g,"")
    .replace(/ヶ/g,"ケ")
    .replace(/之/g,"の");
}


/* =========================
   住所ベース地域判定
========================= */

function classifyAreaByAddress(address){

  const a = normalizeAddressText(address);
  if(!a) return "";


  /* ===== 東京 ===== */

  if(a.includes("江戸川区")){
    if(["西小岩","南小岩","北小岩","上一色","本一色"].some(k=>a.includes(k))) return "江戸川・小岩";
    if(["篠崎","瑞江","一之江","船堀","葛西","西葛西"].some(k=>a.includes(k))) return "江戸川南部";
    return "江戸川";
  }

  if(a.includes("葛飾区")){
    if(["金町","柴又","高砂","新宿"].some(k=>a.includes(k))) return "葛飾東部";
    if(["青戸","立石","亀有","奥戸","堀切"].some(k=>a.includes(k))) return "葛飾西部";
    return "葛飾";
  }

  if(a.includes("足立区")) return "足立";
  if(a.includes("墨田区")) return "墨田";
  if(a.includes("江東区")) return "江東";
  if(a.includes("台東区")) return "台東";


  /* ===== 船橋 ===== */

  if(a.includes("船橋市")){

    if(["西船","西船橋","本中山","二子","二子町","海神","山野町","印内","葛飾町","東中山"].some(k=>a.includes(k))){
      return "船橋西部";
    }

    if(["前原","前原西","前原東","津田沼","東船橋"].some(k=>a.includes(k))){
      return "船橋中部";
    }

    if(["習志野台","薬円台","北習志野","高根台","芝山"].some(k=>a.includes(k))){
      return "船橋東部";
    }

    return "船橋";
  }


  /* ===== 市川 ===== */

  if(a.includes("市川市")){

    if(["八幡","南八幡","鬼越","鬼高","本八幡","東菅野","本北方","北方","若宮","高石神","中山"].some(k=>a.includes(k))){
      return "市川・本八幡";
    }

    if(["中国分","国分","国府台","堀之内","須和田","曽谷","宮久保","菅野"].some(k=>a.includes(k))){
      return "市川北部";
    }

    if(["行徳","妙典","南行徳","相之川","新浜","福栄","末広"].some(k=>a.includes(k))){
      return "行徳・南行徳";
    }

    return "市川";
  }


  /* ===== 松戸 ===== */

  if(a.includes("松戸市")){

    if(["新松戸","新松戸北","新松戸東","幸谷","八ケ崎","八ヶ崎","二ツ木","中和倉"].some(k=>a.includes(k))){
      return "新松戸";
    }

    if(["北松戸","馬橋","西馬橋","栄町","栄町西"].some(k=>a.includes(k))){
      return "北松戸・馬橋";
    }

    if(["常盤平","五香","金ケ作","金ヶ作"].some(k=>a.includes(k))){
      return "常盤平";
    }

    if(["日暮","河原塚","千駄堀","牧の原","松飛台","稔台","みのり台","八柱"].some(k=>a.includes(k))){
      return "八柱・みのり台";
    }

    if(["東松戸","秋山","高塚新田","紙敷"].some(k=>a.includes(k))){
      return "東松戸・秋山";
    }

    if(["北国分","下矢切","上矢切","栗山"].some(k=>a.includes(k))){
      return "北国分・矢切";
    }

    if(["根本","小根本","樋野口","古ケ崎","上本郷"].some(k=>a.includes(k))){
      return "松戸駅周辺";
    }

    return "松戸";
  }


  /* ===== 柏 ===== */

  if(a.includes("柏市")){

    if(["旭町","末広町","柏","中央町"].some(k=>a.includes(k))){
      return "柏駅周辺";
    }

    if(["豊四季","南柏"].some(k=>a.includes(k))){
      return "南柏・豊四季";
    }

    if(["北柏","柏の葉"].some(k=>a.includes(k))){
      return "北柏・柏の葉";
    }

    return "柏";
  }


  /* ===== 流山 ===== */

  if(a.includes("流山市")){

    if(a.includes("南流山")) return "南流山";
    if(a.includes("おおたかの森")) return "流山おおたかの森";

    return "流山";
  }


  if(a.includes("我孫子市")) return "我孫子";
  if(a.includes("鎌ケ谷市") || a.includes("鎌ヶ谷市")) return "鎌ケ谷";


  return "";
}



/* =========================
   座標ベース地域判定
========================= */

function classifyAreaByLatLng(lat,lng){

  if(!isFinite(lat)||!isFinite(lng)) return "";


  /* 東京 */

  if(lng>=139.99){

    if(lat>=35.75) return "葛飾東部";
    if(lat>=35.71) return "江戸川・小岩";

    return "東京東部";
  }


  /* 船橋 */

  if(lat>=35.69 && lat<35.74 && lng>=139.93 && lng<139.99){
    return "船橋西部";
  }


  /* 市川 */

  if(lat>=35.71 && lat<35.76 && lng>=139.89 && lng<139.96){
    return "市川・本八幡";
  }

  if(lat>=35.73 && lat<35.78 && lng>=139.87 && lng<139.93){
    return "市川北部";
  }


  /* 松戸 */

  if(lat>=35.81 && lat<35.86 && lng>=139.90 && lng<139.96){
    return "新松戸";
  }

  if(lat>=35.79 && lat<35.83 && lng>=139.85 && lng<139.91){
    return "常盤平";
  }

  if(lat>=35.76 && lat<35.81 && lng>=139.89 && lng<139.95){
    return "八柱・みのり台";
  }


  /* 柏 */

  if(lat>=35.84 && lng>=139.94 && lng<140.02){
    return "柏";
  }


  /* 流山 */

  if(lat>=35.87 && lng>=139.90 && lng<139.98){
    return "流山";
  }


  return "松戸周辺";
}



/* =========================
   表示地域
========================= */

function getDisplayArea(cast){

  const address=cast.address||"";

  const byAddress=classifyAreaByAddress(address);
  if(byAddress) return byAddress;


  const lat=Number(cast.latitude);
  const lng=Number(cast.longitude);

  if(isFinite(lat)&&isFinite(lng)){
    return classifyAreaByLatLng(lat,lng);
  }

  return cast.area||"";
}
