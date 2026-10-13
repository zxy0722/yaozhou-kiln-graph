/* ============================================================
   耀州窑热词关联图 — 图表与交互逻辑
   数据源：耀州窑研究热点文档(DOCX) × 非遗资源调研摸底登记表(铜川宜君XLSX)
   ============================================================ */
(function () {
  'use strict';

  /* ---------- 读取主题变量 ---------- */
  var style = getComputedStyle(document.documentElement);
  function cssVar(name, fallback) {
    var v = style.getPropertyValue(name).trim();
    return v || fallback;
  }
  var ink    = cssVar('--ink', '#2e2c26');
  var muted  = cssVar('--muted', '#8b8577');
  var rule   = cssVar('--rule', '#e4dfd2');
  var bg2    = cssVar('--bg2', '#ffffff');
  var accent = cssVar('--accent', '#3f6f5f');
  var accent2= cssVar('--accent2', '#a05a2c');

  /* ---------- 门类体系（映射非遗清单10大门类框架） ---------- */
  var CATS = [
    { key: 'core',    name: '窑口核心', color: cssVar('--cat0', '#2f5d4e') },
    { key: 'craft',   name: '技艺工艺', color: cssVar('--cat1', '#4f8a72'), tag: '传统技艺' },
    { key: 'art',     name: '纹饰美学', color: cssVar('--cat2', '#8fa03c'), tag: '传统美术' },
    { key: 'arch',    name: '考古·学术', color: cssVar('--cat3', '#586a9d') },
    { key: 'people',  name: '传承人物', color: cssVar('--cat4', '#b0543f') },
    { key: 'industry',name: '产业融合', color: cssVar('--cat5', '#c08a2e') },
    { key: 'spread',  name: '传播创新', color: cssVar('--cat6', '#7c5d96') }
  ];

  /* ---------- 热词节点（value = 热度权重 1-100） ---------- */
  var RAW_NODES = [
    { name: '耀州窑', cat: 0, value: 100,
      desc: '北方青瓷的代表窑口，窑火延续1400余年；第一批国家级非遗「耀州窑陶瓷烧制技艺」所在地，本关联网络的中心节点。',
      tag: '窑口核心' },

    /* —— 技艺工艺（传统技艺） —— */
    { name: '耀州窑陶瓷烧制技艺', cat: 1, value: 78,
      desc: '第一批国家级非遗代表性项目；孟树锋、陈水平等国家级传承人与张媛等中青年力量构成老中青传承梯队。',
      tag: '传统技艺·国家级' },
    { name: '青瓷', cat: 1, value: 62,
      desc: '耀州窑的核心产品与身份符号，「青如橄榄」的橄榄青釉为宋代鼎盛期典型品种；丝路外销与墓葬考古均以青瓷为实证。',
      tag: '传统技艺' },
    { name: '刻花', cat: 1, value: 50,
      desc: '耀州窑标志性装饰技法，「刀刀见泥」；文化基因研究将「青瓷刻花基因」识别为可量化的核心文化特征。',
      tag: '传统技艺' },
    { name: '剔花', cat: 1, value: 24,
      desc: '与刻花配套的装饰工艺；梁亚萍曾分阶段专项研究黑釉剔花、划花、釉料配方与烧成曲线。',
      tag: '传统技艺' },
    { name: '黑釉', cat: 1, value: 38,
      desc: '唐至清五朝配方、烧成工艺与微观结构持续演变（JECS 2026）；黑褐釉中ε-Fe₂O₃晶体的发现登上国际期刊。',
      tag: '传统技艺' },
    { name: '酱釉', cat: 1, value: 30,
      desc: '铁元素晶相与呈色机理研究对象；2026年在宋代酱釉中发现新型Al₂Li₂FeO₅结构，为潜在反铁磁体候选物。',
      tag: '传统技艺' },
    { name: '橄榄青釉', cat: 1, value: 30,
      desc: '宋代耀州窑典型釉色；「上青下褐」现象成因涉及胎体铁含量、釉层厚度与垫砂工艺（JECS刊文）。',
      tag: '传统技艺' },
    { name: '唐三彩', cat: 1, value: 22,
      desc: '《用ICP-MS元素和TIMS Sr-Nd同位素分析中国唐三彩和耀州窑的特征》（高被引论文）的共同研究对象。',
      tag: '传统技艺' },
    { name: '拉坯', cat: 1, value: 18,
      desc: '核心成型工序；生成式AI数字化采集要求同步记录拉坯、刻花等关键动作的过程影像。',
      tag: '传统技艺' },
    { name: '本土釉料', cat: 1, value: 26,
      desc: '陈水平主持的釉料资料库：累计采集本地原料400余种、梳理1300多组配方、筛选原矿釉800余种，获30余项国家级专利。',
      tag: '传统技艺' },
    { name: '烧成工艺', cat: 1, value: 18,
      desc: '窑炉气氛与烧成时长可调控赤铁矿与ε-Fe₂O₃的生长，是釉色呈色的关键工艺变量。',
      tag: '传统技艺' },
    { name: '窑炉气氛', cat: 1, value: 14,
      desc: '同步辐射原位XRD技术揭示其对釉中铁系晶体生长的调控作用。',
      tag: '传统技艺' },
    { name: '垫砂工艺', cat: 1, value: 12,
      desc: '宋代橄榄青釉器物足部「上青下褐」成因的关键工艺因素之一。',
      tag: '传统技艺' },

    /* —— 纹饰美学（传统美术） —— */
    { name: '文化基因', cat: 2, value: 28,
      desc: '《耀州窑传统陶瓷文化特征识别与基因提取》采用AHP+感性工学语义差异法，提取「青瓷刻花基因」「兰花食器基因」。',
      tag: '传统美术·方法论' },
    { name: '缠枝牡丹', cat: 2, value: 14,
      desc: '孟红娟刻花创作的经典纹饰，耀州窑代表性装饰母题。',
      tag: '传统美术' },
    { name: '双鸭莲花', cat: 2, value: 11,
      desc: '经典刻花纹饰之一，与水波三鱼同属耀州窑传统纹样谱系。',
      tag: '传统美术' },
    { name: '老碗', cat: 2, value: 12,
      desc: '张媛代表作，以十二个主题呈现关中民俗，是陶瓷艺术与民俗生活的结合。',
      tag: '传统美术·作品' },

    /* —— 考古·学术 —— */
    { name: '漆水河流域', cat: 3, value: 20,
      desc: '耀州窑文明根系所在，追溯至仰韶文化时期，揭示「火与土之歌」的史前序章。',
      tag: '文明溯源' },
    { name: '仰韶文化', cat: 3, value: 20,
      desc: '约6000年前「人首龙身」贴塑纹陶罐出土于此，比红山文化「玉猪龙」更直接地链接神话谱系。',
      tag: '史前考古' },
    { name: '人首龙身', cat: 3, value: 16,
      desc: '贴塑纹陶罐纹样，直接链接伏羲、女娲神话谱系，被视为耀州窑史前文化基因。',
      tag: '史前纹样' },
    { name: '杜回村墓葬', cat: 3, value: 24,
      desc: '西安长安区北宋墓葬出土60件耀州窑精品青瓷，墓主孟氏家族有成员在耀州为官，为耀州瓷流通与消费提供实证。',
      tag: '墓葬考古' },
    { name: '狻猊钮盖熏炉', cat: 3, value: 13,
      desc: '杜回村北宋墓出土青瓷代表器型，与凤头执壶同为北宋耀州瓷断代实证。',
      tag: '器物研究' },
    { name: '苏峪口瓷窑址', cat: 3, value: 12,
      desc: '2025年度中国十大考古新发现，与耀州窑同属北方瓷业体系，是各民族文化交融互鉴的例证。',
      tag: '窑址考古' },
    { name: '辽墓辽塔', cat: 3, value: 12,
      desc: 'Springer辽代陶瓷研究专著指出：辽墓与辽塔出土耀州器物为五代至北宋耀州瓷提供可靠断代依据。',
      tag: '断代研究' },
    { name: '断代标尺', cat: 3, value: 12,
      desc: '耀州瓷在辽墓与辽塔出土器物中形成的年代判定基准，凸显耀州窑的编年价值。',
      tag: '考古方法' },
    { name: 'ε-Fe₂O₃晶体', cat: 3, value: 28,
      desc: 'Journal of Physical Chemistry C 揭示黑褐釉中该晶体的热行为，证明窑炉气氛和烧成时长可调控其生长。',
      tag: '材料科学' },
    { name: '反铁磁体', cat: 3, value: 20,
      desc: '宋代耀州窑酱釉中发现新型Al₂Li₂FeO₅结构，被识别为潜在反铁磁体候选物，是科技考古的跨学科突破。',
      tag: '凝聚态物理' },
    { name: '呈色机理', cat: 3, value: 24,
      desc: '酱釉、黑釉中铁元素晶相与呈色机理研究，是耀州窑科技分析的核心方向之一。',
      tag: '科技分析' },
    { name: 'ICP-MS同位素分析', cat: 3, value: 22,
      desc: '《用ICP-MS元素和TIMS Sr-Nd同位素分析中国唐三彩和耀州窑的特征》为知网高被引论文。',
      tag: '科技分析' },
    { name: '文献计量分析', cat: 3, value: 16,
      desc: '《基于文献计量分析的1949—2023年中国耀州瓷研究》以VOSviewer揭示：热点已从烧制考古转向技艺传承与文化传播。',
      tag: '研究综述' },

    /* —— 传承人物 —— */
    { name: '孟树锋', cat: 4, value: 36,
      desc: '国家级非遗代表性传承人；1976年与李国桢历时三年复原断代800年的耀州青瓷，参与「瓷话千年·匠新对话」跨界对话。',
      tag: '国家级传承人' },
    { name: '陈水平', cat: 4, value: 30,
      desc: '第一批国家级非遗代表性传承人、陕西省工艺美术大师；建立本土釉料产业生态，带徒原则「可以大胆创作，但不能丢掉耀州窑文化根基」。',
      tag: '国家级传承人' },
    { name: '张媛', cat: 4, value: 28,
      desc: '全国五一劳动奖章获得者（2026）、市级代表性传承人；编写省内高职首个陶瓷人才培养方案，构建「四位一体」模式。',
      tag: '市级传承人' },
    { name: '袁明', cat: 4, value: 28,
      desc: '电商破局者；2024年入驻耀州窑直播基地，与「与辉同行」合作实现10万套「平安喜乐杯」秒空，全年线上销售额突破800万元。',
      tag: '产业新锐' },
    { name: '孟红娟', cat: 4, value: 20,
      desc: '陕西省工艺美术大师，在耀州窑非遗刻花展示区从事刻花创作，手执刻刀刻画经典纹饰。',
      tag: '省工艺美术大师' },
    { name: '梁亚萍', cat: 4, value: 22,
      desc: '唐宋陶业艺术总监；2021年成立全国三八红旗手工作站，从装饰工艺起步系统研究黑釉剔花、划花、釉料配方与烧成曲线。',
      tag: '系统传承者' },
    { name: '孟鸣', cat: 4, value: 18,
      desc: '中国孟梓窑第六代嫡传人，师从其父孟树锋；实行三级产品体系，引入3D打印打样与AI设计工具。',
      tag: '新生代传承' },
    { name: '徐军照', cat: 4, value: 13,
      desc: '在2026西安丝绸之路国际旅游博览会设置陶艺体验装置，手把手指导市民体验揉泥、拉坯。',
      tag: '体验推广' },
    { name: '乡村工匠', cat: 4, value: 14,
      desc: '茹春玲、王燕利入选2025年省级乡村工匠名师；王家砭村陶瓷文化产业园带动近200名村民就近就业。',
      tag: '乡村振兴' },

    /* —— 产业融合 —— */
    { name: '直播电商', cat: 5, value: 30,
      desc: '耀州窑直播基地联动「与辉同行」等平台，推动耀瓷销售从「以产定销」转向「以销定产」。',
      tag: '产业新渠道' },
    { name: '以销定产', cat: 5, value: 18,
      desc: '袁明提出的产销新逻辑，重塑耀瓷生产组织方式。',
      tag: '产业模式' },
    { name: '平安喜乐杯', cat: 5, value: 16,
      desc: '与「与辉同行」合作的爆款产品，10万套上线即秒空。',
      tag: '爆款案例' },
    { name: '陶瓷+餐饮', cat: 5, value: 15,
      desc: '铜川市烹饪餐饮行业协会组织10余家餐饮企业集中采购唐宋陶业特色餐具，耀州瓷从展厅走向百姓餐桌。',
      tag: '跨界案例' },
    { name: '工业陶瓷', cat: 5, value: 20,
      desc: '耀州窑文化基地近三年产业链规模年均增长14.29%，落地8家工业陶瓷企业，4家培育为规上企业。',
      tag: '产业升级' },
    { name: '碳化硅', cat: 5, value: 12,
      desc: '碳化硅陶瓷项目达产后预计年产高性能制品800吨以上，年产值1.2亿元。',
      tag: '新材料' },
    { name: '氧化铝', cat: 5, value: 12,
      desc: '技改完成后年产新能源汽车特种陶瓷材料300吨、高压熔断器精密瓷壳60万只，预计年新增产值3000万元。',
      tag: '新材料' },
    { name: '职业教育', cat: 5, value: 16,
      desc: '铜川职院陶瓷设计与工艺专业建成全国陶瓷行业技能评价基地，与10多家企业合作育才。',
      tag: '人才培养' },
    { name: '非遗工坊', cat: 5, value: 18,
      desc: '清单显示宜君县已认定3家省级非遗工坊；「工坊模式」是非遗「带动传承+产业造血」的核心载体。',
      tag: '清单数据' },

    /* —— 传播创新 —— */
    { name: '丝路贸易', cat: 6, value: 28,
      desc: '《耀州青瓷与北宋陆上丝路延续性的互证研究》采用三重复合论证模型，实证耀州青瓷的丝路角色。',
      tag: '丝路研究' },
    { name: '官退民进', cat: 6, value: 18,
      desc: '《浙江工商大学学报》2025年刊文揭示北宋陆上丝路「官退民进」的商贸转型特征。',
      tag: '丝路研究' },
    { name: '汉服联动', cat: 6, value: 22,
      desc: '第三届汉服设计大赛暨耀州瓷非遗焕新季设三大赛道，以双IP联动搭建产学研平台。',
      tag: '跨界IP' },
    { name: '耀州窑博物馆', cat: 6, value: 32,
      desc: '2026年7月佛得角、拉脱维亚、博茨瓦纳等9国驻华外交官到访体验古法制瓷全流程；知网发文量第二的研究机构。',
      tag: '文化空间' },
    { name: '驻华外交官', cat: 6, value: 15,
      desc: '9国驻华外交官沉浸式体验古法制瓷，拉脱维亚大使盛赞「中国古人高雅的审美」。',
      tag: '国际传播' },
    { name: '研学体验', cat: 6, value: 18,
      desc: '陶艺体验、非遗课堂、古法制瓷体验等构成「可参与」的非遗传播场景。',
      tag: '体验传播' },
    { name: '生成式AI', cat: 6, value: 32,
      desc: '《生成式AI赋能耀州窑非遗数字化保护与设计创新研究》提出「可控生成」框架：以工艺可行性和文化语义为硬约束。',
      tag: '数字创新' },
    { name: '数字化保护', cat: 6, value: 24,
      desc: '强调同步记录釉色多角度光照表现与关键动作过程影像，构建多模态数字档案。',
      tag: '数字创新' }
  ];

  /* ---------- 关联边（[源, 目标, 关联强度1-5]） ---------- */
  var RAW_EDGES = [
    /* 中心辐射 */
    ['耀州窑', '耀州窑陶瓷烧制技艺', 5], ['耀州窑', '青瓷', 5], ['耀州窑', '刻花', 4],
    ['耀州窑', '耀州窑博物馆', 4], ['耀州窑', '孟树锋', 4], ['耀州窑', '漆水河流域', 3],
    ['耀州窑', '丝路贸易', 4], ['耀州窑', '直播电商', 3], ['耀州窑', '生成式AI', 4],
    ['耀州窑', '杜回村墓葬', 3], ['耀州窑', '工业陶瓷', 3], ['耀州窑', '文献计量分析', 3],
    ['耀州窑', '苏峪口瓷窑址', 2], ['耀州窑', 'ICP-MS同位素分析', 2],
    /* 技艺脉络 */
    ['青瓷', '刻花', 5], ['青瓷', '橄榄青釉', 4], ['刻花', '剔花', 2],
    ['黑釉', '呈色机理', 4], ['黑釉', 'ε-Fe₂O₃晶体', 4], ['黑釉', '烧成工艺', 3],
    ['烧成工艺', '窑炉气氛', 3], ['窑炉气氛', 'ε-Fe₂O₃晶体', 3], ['酱釉', '反铁磁体', 4],
    ['酱釉', '呈色机理', 3], ['橄榄青釉', '垫砂工艺', 2], ['唐三彩', 'ICP-MS同位素分析', 4],
    ['耀州窑陶瓷烧制技艺', '本土釉料', 3], ['耀州窑陶瓷烧制技艺', '拉坯', 2],
    /* 人物传承 */
    ['耀州窑陶瓷烧制技艺', '孟树锋', 5], ['耀州窑陶瓷烧制技艺', '陈水平', 5],
    ['耀州窑陶瓷烧制技艺', '张媛', 4], ['本土釉料', '陈水平', 4],
    ['张媛', '职业教育', 4], ['张媛', '老碗', 2], ['张媛', '乡村工匠', 2],
    ['孟树锋', '孟鸣', 4], ['孟树锋', '汉服联动', 3], ['孟树锋', '青瓷', 2],
    ['孟红娟', '刻花', 4], ['孟红娟', '缠枝牡丹', 3], ['孟红娟', '双鸭莲花', 2],
    ['梁亚萍', '剔花', 3], ['梁亚萍', '黑釉', 3], ['梁亚萍', '陶瓷+餐饮', 3],
    ['袁明', '直播电商', 5], ['袁明', '以销定产', 3], ['袁明', '平安喜乐杯', 3],
    ['徐军照', '研学体验', 3], ['徐军照', '拉坯', 2], ['孟鸣', '生成式AI', 3],
    /* 考古实证 */
    ['漆水河流域', '仰韶文化', 4], ['仰韶文化', '人首龙身', 3],
    ['杜回村墓葬', '狻猊钮盖熏炉', 3], ['杜回村墓葬', '青瓷', 4],
    ['辽墓辽塔', '断代标尺', 3], ['断代标尺', '青瓷', 3],
    /* 产业脉络 */
    ['直播电商', '以销定产', 3], ['工业陶瓷', '碳化硅', 3], ['工业陶瓷', '氧化铝', 3],
    ['非遗工坊', '研学体验', 2], ['非遗工坊', '耀州窑陶瓷烧制技艺', 2],
    /* 传播脉络 */
    ['丝路贸易', '官退民进', 3], ['丝路贸易', '青瓷', 4],
    ['耀州窑博物馆', '驻华外交官', 3], ['耀州窑博物馆', '研学体验', 3],
    ['生成式AI', '数字化保护', 4], ['生成式AI', '刻花', 3], ['生成式AI', '拉坯', 2],
    ['数字化保护', '文化基因', 2], ['文化基因', '刻花', 3]
  ];

  /* ---------- 工具函数 ---------- */
  function nodeSize(v) {
    if (v >= 100) return 74;
    return Math.max(17, 10 + v * 0.52);
  }

  var NODE_INDEX = {};
  RAW_NODES.forEach(function (n, i) { NODE_INDEX[n.name] = i; });

  var graphDom = document.getElementById('chart-graph');
  if (graphDom) {
    var chart = echarts.init(graphDom, null, { renderer: 'canvas' });
    var activeCats = CATS.map(function (_, i) { return i; }); // 全部可见
    var selectedIdx = -1;

    function buildGraphData() {
      var nodes = RAW_NODES.map(function (n) {
        var c = CATS[n.cat];
        return {
          name: n.name,
          value: n.value,
          category: n.cat,
          symbolSize: nodeSize(n.value),
          itemStyle: { color: c.color, borderColor: bg2, borderWidth: 1.5 },
          label: {
            show: true,
            fontSize: n.value >= 100 ? 15 : (n.value >= 30 ? 12 : 10.5),
            fontWeight: n.value >= 60 ? 700 : 400,
            color: n.value >= 100 ? '#ffffff' : ink
          },
          desc: n.desc,
          tag: n.tag
        };
      });
      var edges = RAW_EDGES
        .filter(function (e) {
          return activeCats.indexOf(RAW_NODES[NODE_INDEX[e[0]]].cat) >= 0 &&
                 activeCats.indexOf(RAW_NODES[NODE_INDEX[e[1]]].cat) >= 0;
        })
        .map(function (e) {
          return {
            source: e[0], target: e[1], weight: e[2],
            lineStyle: { width: 0.6 + e[2] * 0.55, opacity: 0.32 + e[2] * 0.06 }
          };
        });
      return { nodes: nodes, edges: edges };
    }

    function graphOption() {
      var data = buildGraphData();
      return {
        animationDuration: 1200,
        animationEasingUpdate: 'quinticInOut',
        tooltip: {
          appendToBody: true,
          backgroundColor: 'rgba(255,255,255,0.96)',
          borderColor: rule,
          textStyle: { color: ink, fontSize: 12.5 },
          extraCssText: 'max-width:340px;white-space:normal;box-shadow:0 6px 24px rgba(46,44,38,.12);border-radius:10px;padding:10px 14px;',
          formatter: function (p) {
            if (p.dataType === 'edge') return p.data.source + ' ⇄ ' + p.data.target;
            var n = RAW_NODES[p.dataIndex];
            return '<b style="font-size:14px">' + n.name + '</b>' +
              '<div style="margin:4px 0 6px"><span style="display:inline-block;padding:1px 8px;border-radius:99px;font-size:11px;color:#fff;background:' +
              CATS[n.cat].color + '">' + n.tag + '</span>' +
              '<span style="margin-left:6px;font-size:11px;color:' + muted + '">热度 ' + n.value + '</span></div>' +
              '<div style="line-height:1.65;color:' + ink + '">' + n.desc + '</div>';
          }
        },
        legend: { show: false },
        series: [{
          type: 'graph',
          layout: 'force',
          roam: true,
          draggable: true,
          data: data.nodes,
          links: data.edges,
          categories: CATS.map(function (c) { return { name: c.name }; }),
          force: {
            repulsion: 300,
            gravity: 0.14,
            edgeLength: [55, 150],
            friction: 0.18,
            layoutAnimation: true
          },
          label: { position: 'bottom', distance: 4 },
          itemStyle: { shadowBlur: 8, shadowColor: 'rgba(46,44,38,0.18)' },
          lineStyle: { color: 'source', curveness: 0.16 },
          emphasis: {
            focus: 'adjacency',
            lineStyle: { opacity: 0.85 },
            itemStyle: { shadowBlur: 18, shadowColor: 'rgba(63,111,95,0.4)' }
          },
          scaleLimit: { min: 0.5, max: 3 }
        }]
      };
    }

    chart.setOption(graphOption());

    /* 节点点击 → 详情面板 */
    chart.on('click', function (p) {
      if (p.dataType !== 'node') return;
      showDetail(p.dataIndex);
    });
    chart.getZr().on('click', function (e) {
      if (!e.target) { selectedIdx = -1; hideDetail(); }
    });

    window.addEventListener('resize', function () { chart.resize(); });

    /* ---------- 门类筛选 chips ---------- */
    var chipWrap = document.getElementById('cat-chips');
    var counts = {};
    RAW_NODES.forEach(function (n) { counts[n.cat] = (counts[n.cat] || 0) + 1; });
    CATS.forEach(function (c, i) {
      var b = document.createElement('button');
      b.className = 'chip on';
      b.style.setProperty('--chip', c.color);
      b.innerHTML = '<i></i>' + c.name + ' <em>' + (counts[i] || 0) + '</em>';
      b.addEventListener('click', function () {
        var pos = activeCats.indexOf(i);
        if (pos >= 0 && activeCats.length > 1) { activeCats.splice(pos, 1); b.classList.remove('on'); }
        else if (pos < 0) { activeCats.push(i); b.classList.add('on'); }
        chart.setOption(graphOption());
        selectedIdx = -1; hideDetail();
      });
      chipWrap.appendChild(b);
    });

    /* ---------- 搜索定位 ---------- */
    var searchInput = document.getElementById('kw-search');
    searchInput.addEventListener('change', function () {
      var q = searchInput.value.trim();
      var idx = -1;
      RAW_NODES.forEach(function (n, i) { if (n.name.indexOf(q) >= 0 && idx < 0) idx = i; });
      if (q && idx >= 0) {
        chart.dispatchAction({ type: 'focusNodeAdjacency', seriesIndex: 0, dataIndex: idx });
        chart.dispatchAction({ type: 'showTip', seriesIndex: 0, dataIndex: idx });
        showDetail(idx);
      } else {
        chart.dispatchAction({ type: 'unfocusNodeAdjacency', seriesIndex: 0 });
        chart.dispatchAction({ type: 'downplay', seriesIndex: 0 });
        chart.dispatchAction({ type: 'hideTip' });
      }
    });
    document.getElementById('btn-reset').addEventListener('click', function () {
      activeCats = CATS.map(function (_, i) { return i; });
      Array.prototype.forEach.call(chipWrap.children, function (b) { b.classList.add('on'); });
      searchInput.value = '';
      selectedIdx = -1; hideDetail();
      chart.setOption(graphOption(), true);
    });

    /* ---------- 详情面板 ---------- */
    var panel = document.getElementById('detail-panel');
    function showDetail(idx) {
      selectedIdx = idx;
      var n = RAW_NODES[idx];
      var related = [];
      RAW_EDGES.forEach(function (e) {
        if (e[0] === n.name) related.push(e[1]);
        else if (e[1] === n.name) related.push(e[0]);
      });
      var relHtml = related.map(function (r) {
        var rn = RAW_NODES[NODE_INDEX[r]];
        return '<button class="rel-tag" data-name="' + r + '" style="--chip:' + CATS[rn.cat].color + '">' + r + '</button>';
      }).join('');
      panel.innerHTML =
        '<div class="d-head"><span class="d-dot" style="background:' + CATS[n.cat].color + '"></span>' +
        '<div><h4>' + n.name + '</h4><span class="d-meta">' + n.tag + ' · 热度 ' + n.value + '</span></div></div>' +
        '<p class="d-desc">' + n.desc + '</p>' +
        '<div class="d-sub">关联热词（' + related.length + '）</div>' +
        '<div class="d-rel">' + relHtml + '</div>';
      panel.classList.add('has');
      Array.prototype.forEach.call(panel.querySelectorAll('.rel-tag'), function (btn) {
        btn.addEventListener('click', function () {
          var i = NODE_INDEX[btn.getAttribute('data-name')];
          chart.dispatchAction({ type: 'focusNodeAdjacency', seriesIndex: 0, dataIndex: i });
          chart.dispatchAction({ type: 'showTip', seriesIndex: 0, dataIndex: i });
          showDetail(i);
        });
      });
    }
    function hideDetail() {
      panel.classList.remove('has');
      panel.innerHTML = '<div class="d-empty">点击图中热词节点，<br>查看词条解读与关联网络</div>';
      chart.dispatchAction({ type: 'unfocusNodeAdjacency', seriesIndex: 0 });
      chart.dispatchAction({ type: 'downplay', seriesIndex: 0 });
    }
  }

  /* ============================================================
     图2 非遗清单门类分布（数据：登记表·宜君县，n=52）
     ============================================================ */
  var catDom = document.getElementById('chart-cats');
  if (catDom) {
    var catChart = echarts.init(catDom, null, { renderer: 'svg' });
    var catData = [
      { name: '传统技艺', v: 23, c: cssVar('--cat1', '#4f8a72') },
      { name: '民间文学', v: 10, c: cssVar('--cat3', '#586a9d') },
      { name: '传统舞蹈', v: 10, c: cssVar('--cat6', '#7c5d96') },
      { name: '传统美术', v: 7, c: cssVar('--cat2', '#8fa03c') },
      { name: '传统音乐', v: 1, c: cssVar('--cat5', '#c08a2e') },
      { name: '民俗', v: 1, c: cssVar('--cat4', '#b0543f') }
    ];
    catChart.setOption({
      animation: false,
      grid: { left: 8, right: 40, top: 8, bottom: 0, containLabel: true },
      tooltip: {
        appendToBody: true,
        trigger: 'axis',
        axisPointer: { type: 'none' },
        backgroundColor: 'rgba(255,255,255,0.96)', borderColor: rule,
        textStyle: { color: ink, fontSize: 12.5 },
        formatter: function (ps) {
          var p = ps[0];
          var total = 52;
          return '<b>' + p.name + '</b><br>' + p.value + ' 项 · 占比 ' + (p.value / total * 100).toFixed(1) + '%';
        }
      },
      xAxis: { type: 'value', show: false },
      yAxis: {
        type: 'category', inverse: true,
        data: catData.map(function (d) { return d.name; }),
        axisLine: { show: false }, axisTick: { show: false },
        axisLabel: { color: ink, fontSize: 13 }
      },
      series: [{
        type: 'bar', barWidth: 18,
        data: catData.map(function (d) { return { value: d.v, itemStyle: { color: d.c, borderRadius: [0, 9, 9, 0] } }; }),
        label: { show: true, position: 'right', color: muted, fontSize: 12.5, fontWeight: 600, formatter: '{c} 项' }
      }]
    });
    window.addEventListener('resize', function () { catChart.resize(); });
  }

  /* ============================================================
     图3 热词热度排行 Top 15
     ============================================================ */
  var rankDom = document.getElementById('chart-rank');
  if (rankDom) {
    var rankChart = echarts.init(rankDom, null, { renderer: 'svg' });
    var top = RAW_NODES.slice().sort(function (a, b) { return b.value - a.value; }).slice(0, 15).reverse();
    rankChart.setOption({
      animation: false,
      grid: { left: 8, right: 52, top: 8, bottom: 0, containLabel: true },
      tooltip: {
        appendToBody: true, trigger: 'axis', axisPointer: { type: 'none' },
        backgroundColor: 'rgba(255,255,255,0.96)', borderColor: rule,
        textStyle: { color: ink, fontSize: 12.5 },
        formatter: function (ps) {
          var p = ps[0];
          var n = RAW_NODES[NODE_INDEX[p.name]];
          return '<b>' + p.name + '</b> <span style="color:' + muted + '">· ' + n.tag + '</span><br>热度权重 ' + p.value;
        }
      },
      xAxis: { type: 'value', show: false, max: 100 },
      yAxis: {
        type: 'category',
        data: top.map(function (n) { return n.name; }),
        axisLine: { show: false }, axisTick: { show: false },
        axisLabel: { color: ink, fontSize: 12.5 }
      },
      series: [{
        type: 'bar', barWidth: 16,
        data: top.map(function (n) {
          return { value: n.value, itemStyle: { borderRadius: [0, 8, 8, 0] } };
        }),
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: accent }, { offset: 1, color: accent2 }
          ])
        },
        label: { show: true, position: 'right', color: muted, fontSize: 12, fontWeight: 600 }
      }]
    });
    window.addEventListener('resize', function () { rankChart.resize(); });
  }
})();
