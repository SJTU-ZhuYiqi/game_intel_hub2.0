const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, LevelFormat } = require('docx');
const fs = require('fs');

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 36, color: '0f1f3d', font: '宋体' })],
    spacing: { before: 560, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: '1a2f52', space: 4 } },
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 28, color: '1a2f52', font: '宋体' })],
    spacing: { before: 400, after: 100 },
  });
}
function b1(text) {
  return new Paragraph({
    numbering: { reference: 'b1', level: 0 },
    children: [new TextRun({ text, size: 22, font: '宋体', color: '0f172a' })],
    spacing: { before: 120, after: 40, line: 340, lineRule: 'auto' },
  });
}
function b2(text) {
  return new Paragraph({
    numbering: { reference: 'b2', level: 0 },
    children: [new TextRun({ text, size: 21, font: '宋体', color: '334155' })],
    spacing: { before: 60, after: 40, line: 320, lineRule: 'auto' },
  });
}
function b2b(bold, rest) {
  return new Paragraph({
    numbering: { reference: 'b2', level: 0 },
    children: [
      new TextRun({ text: bold, bold: true, size: 21, font: '宋体', color: '1e3a6a' }),
      new TextRun({ text: rest, size: 21, font: '宋体', color: '334155' }),
    ],
    spacing: { before: 60, after: 40, line: 320, lineRule: 'auto' },
  });
}
function div() {
  return new Paragraph({
    children: [new TextRun('')],
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'e2e8f0' } },
    spacing: { before: 160, after: 160 },
  });
}
function sp() {
  return new Paragraph({ children: [new TextRun('')], spacing: { before: 80, after: 80 } });
}

const numbering = {
  config: [
    { reference: 'b1', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u25CF', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 280 } } } }] },
    { reference: 'b2', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 880, hanging: 280 } } } }] },
  ],
};

const children = [
  new Paragraph({
    children: [new TextRun({ text: 'W17 内容消费行业资讯周报', bold: true, size: 56, color: '0f1f3d', font: '宋体' })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 1440, after: 240 },
  }),
  new Paragraph({
    children: [new TextRun({ text: '2026 年第 17 周（4 月 27 日 ~ 5 月 5 日）', size: 26, color: '475569', font: '宋体' })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
  }),
  new Paragraph({
    children: [new TextRun({ text: '资讯快评 · 对业务意味着什么 | 内容消费研究组内部参考', size: 22, color: '94a3b8', font: '宋体' })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 2400 },
  }),

  h1('一、游戏行业 | 资讯快评——对业务意味着什么'),

  // G1 掌上谈兵持续霸榜
  h2('掌上谈兵持续霸榜，疯狂水世界6天称霸，小游戏双雄格局确立'),
  b1('连续两周霸榜不是偶然——下沉流量的小游戏分流在五一期间正式从"趋势"变成"结论"，快手不跟进就是在拱手让出'),
  b2('掌上谈兵上线第二周仍居榜首，首周的"大作哑火+时机窗口"逻辑已经失效，这款产品本身的留存和变现结构支撑了它的持续霸榜；对快手来说，这意味着如果买量合作尚未落地，已经错过了产品最高热度窗口，成本窗口在缩窄'),
  b2('疯狂水世界连续6天榜首，且抖音小游戏同期也保持前列；双端买量意味着益世界发行团队预算充沛、触达多平台意愿强——快手能拿到多少份额，取决于能不能给出差异化的用户池价值（下沉+主播导量），而不是靠价格竞争'),
  b2b('数据侧追问：', '快手平台上掌上谈兵和疯狂水世界的当前消耗量级是多少？如果还处于小体量测试阶段，需要找到是"产品方不愿意投"还是"快手用户池转化率不达标"，两个原因对应完全不同的应对策略'),
  b2('超能下蛋鸭登顶抖音小游戏——IAA轻度品类在五一假期的爆发是可预期的规律；快手的IAA小游戏生态是否同步出现类似产品？假期结束后IAA轻度品的留存通常断崖，下周数据更有参考价值'),
  sp(),

  // G2 五一档二游年庆
  h2('崩铁三周年挤进iOS TOP5，鸣潮一周年双城开花，五一档二游买量进入高烈度期'),
  b1('二游年庆买量竞争烈度在五一达到上半年峰值，但对快手的含义是"防守不失血"而不是"进攻拿份额"'),
  b2('崩铁上海线下近3万人参与，线下活动把核心用户的注意力和社交资产都绑定在米哈游生态里，这批用户在五一期间被极大分心，外部平台（包括快手）对他们的触达效率在这个窗口显著下降；二游买量CPM五一期间大概率走高，这一段不是扩量的好时机'),
  b2('鸣潮一周年库洛在杭州+广州双城线下；库洛是快手买量中相对活跃的二游广告主之一，年庆后的回调期（5月中下旬）才是争差价窗口——年庆期预算主要打自有生态，年庆结束后才会向外买量平台溢出'),
  b2b('可以建立的观测：', '拉一下崩铁/鸣潮在快手平台的历史消耗曲线，看年庆前后各有多少波动幅度，测算今年年庆结束后的回调窗口时间点，提前一周找客户提预算档期'),
  sp(),

  // G3 韩国OneStore微信小游戏出海
  h2('韩国OneStore宣布5月引进微信小游戏，ARPPU为Google Play的5倍'),
  b1('小游戏出海渠道新窗口开启，意义不只是"又多了个渠道"，而是ARPPU溢价验证了下沉小游戏品类的海外高净值用户机会'),
  b2('OneStore是SKT/KT/LG三大运营商+Naver的联合体，覆盖韩国主流手机预装，ARPPU是Google Play的5倍说明这批用户付费意愿强、游戏品类重合度高（SLG/卡牌/MMO）；这对正在出海的中国小游戏厂商来说是正向信号'),
  b2('对快手的含义有两层：一是快手合作的小游戏广告主（掌上谈兵、疯狂水世界、无尽冬日等）如果在出海布局，OneStore是优先级较高的渠道，快手能不能在国内买量合作基础上延伸到出海协同？二是快手国际化业务（Kwai）的用户结构是否与这类小游戏的海外目标用户有重叠？'),
  b2b('需要核实的信息：', '目前快手有没有面向出海小游戏厂商的广告产品，能否在国内合作中附带提一下出海协同方案？如果有，这是差异化销售话术；如果没有，这是产品空白'),
  sp(),

  // G4 异环一周复盘
  h2('《异环》上线一周：4天iOS收入1163万，官方致歉，移动端优化仍是核心短板'),
  b1('异环的问题从"没达到预期"变成"主动承认失败并承诺修复"，预算决策变得更不确定，快手跟进节奏需要和外部修复节点绑定'),
  b2('4月29日官方发致歉信并承诺修复移动端优化问题，说明完美世界高管层已承认移动端是硬伤；但修复周期通常是数周到数月，这段时间内买量ROI不可信——即使单日素材量升至7700条，素材同质化严重、转化率低，加大投放只是在稀释预算'),
  b2('对快手销售侧的含义：完美世界当前的买量行为（7700条素材）更像是"用投放量撑声量"而不是真正在拿ROI，这种买量会集中在优化成本低的平台，快手未必是优先级；与其被动等，不如明确一个重新评估节点——5月中海外版上线数据，或修复版本上线后的次周留存，再决定是否主动推进'),
  b2b('竞品时间线：', '网易《无限大》和米哈游《Varsapura》预计2026下半年~2027年上线，这意味着异环如果在这之前稳定下来，仍有扩张窗口；当前是窗口期，但需要条件触发'),
  sp(),

  // G5 GTA6
  h2('GTA6确认11月16日主机独占发售，首日预测2000万份'),
  b1('GTA6是全年最大的行业事件，但对快手买量业务几乎没有直接贡献——反而是观察移动端流量迁移的机会'),
  b2('GTA6主机独占意味着完全绕开移动端和PC买量生态，快手没有直接参与买量的可能；但可以观察11月GTA6发售前后，移动端游戏买量消耗是否受到压制（玩家注意力转移到主机），这是个自然实验，对预测Q4游戏买量大盘有参考价值'),
  b2b('对广告主报告时的注意事项：', 'GTA6对行业的拉动主要体现在PC+主机端，如果广告主用GTA6预期支撑移动端买量预算的乐观判断，需要提醒这个拉动逻辑不传导'),
  div(),

  new Paragraph({ children: [new TextRun('')], pageBreakBefore: true }),
  h1('二、短剧漫剧 | 资讯快评——对业务意味着什么'),

  // D1 菩提临世下架
  h2('《菩提临世》8亿播放下架：爆款寿命28天，合规边界正式拉齐真人短剧'),
  b1('AI漫剧爆款的可持续性问题从"潜在风险"变成了"已发生事实"，对快手漫剧广告业务意味着供给质量评估需要加入合规维度，不能只看播放量'),
  b2('菩提临世8亿播放是真实数据，但下架说明"流量大"和"可持续变现"之间有一道合规门槛；对快手来说，如果漫剧广告库存里有类似灵异/宗教/复仇类题材产品，需要主动评估其备案状态，避免广告主在CP下架后出现品牌安全投诉'),
  b2('下架的直接受益者是《彪悍人生》——不是因为彪悍人生更好，而是因为它在合规赛道上没有对手；这提示了一个可量化的问题：快手平台上当前在播的AI漫剧里，完成广电备案的占比是多少？这个比例决定了快手能向品牌广告主提供的"合规库存"有多大'),
  b2b('广告主沟通时的注意点：', '如果广告主问"快手的漫剧内容安全吗"，需要有一个明确的合规CP白名单和备案状态核查流程，而不是只说"我们会审核"；菩提临世事件后，品牌广告主对这个问题的敏感度显著上升'),
  sp(),

  // D2 红果VIP收费风波
  h2('红果VIP收费冲上热搜（阅读量2.4亿），官方辟谣：免费模式未变'),
  b1('表面上是辟谣，本质上是竞品商业化能力在提升，快手短剧的免费定位需要重新审视是否仍然是护城河'),
  b2('红果月活已达3亿（+84.6%），人均使用时长125分钟/天，这个量级意味着它不再是抖音附属产品，而是独立的内容消费入口；抖音4月新设"红果电商"独立部门说明抖音在推进"内容即货架"的商业化闭环，短剧免费模式背后的商业化逻辑已经比纯广告分账复杂得多'),
  b2('对快手的含义：快手短剧目前的商业化主要是广告分账，如果红果把VIP订阅+电商带货+广告三条线都跑通，快手的单一广告模式在ARPU上处于劣势；这不是本周的紧急问题，但是一个值得数据侧追踪的中长期信号——可以持续观察红果月活增速和广告主向红果迁移的预算比例'),
  b2b('可以向运营侧提的问题：', '快手短剧内容在快手站内的流量分配逻辑，和红果的漫剧首发合作深度相比，谁给CP的变现条件更好？如果红果给的更好，头部CP会优先给红果独家首发，快手只能拿二轮'),
  sp(),

  // D3 清朗专项+广电备案，监管三重成型
  h2('网信办清朗专项+广电备案新规+平台自审，AI漫剧监管三重框架4月正式成型'),
  b1('监管框架成型是行业结构性变量，不是阶段性事件；对快手漫剧业务，这意味着合规能力本身将成为差异化的销售话语权'),
  b2('三重合力的实际效果是：小CP的试错成本大幅上升，低质供给的出清速度加快；这对快手短期广告库存有压缩效果（合规内容方数量少），但中长期对提升广告主品牌安全感有利——这是一个"先痛后利"的结构'),
  b2('广电备案新规（4月1日起）+网信清朗专项（4月30日）给了平台一个主动梳理库存的理由；快手运营侧是否已经启动了一轮AI漫剧供给的合规审查？如果没有，在广告主侧主动推广合规内容之前，需要先把自己的底仓清一遍'),
  b2b('对销售侧的正向价值：', '当广告主问"为什么投快手漫剧比投抖音/红果更安全"，监管框架成型后，快手如果有一套透明的合规CP准入标准和备案核查流程，这本身就是销售话术；建议梳理成一页可以对外展示的合规能力说明'),
  sp(),

  // D4 掌阅泡漫工业化跑通
  h2('掌阅泡漫《彪悍人生》5.8亿，短剧营收+139%，毛利率68%——AI短剧工业化路径被验证'),
  b1('掌阅不只是验证了一部剧的成功，而是验证了"IP版权+AI流水线+双平台分发+自动化投放"的完整商业闭环，这个模式一旦被复制，行业头部化进程会加快'),
  b2('掌阅的竞争优势来自三个叠加：拥有大量网文IP版权（省去改编授权成本）、泡漫平台实现了"剧本进+分镜出"的AI全流程（压低制作成本）、自研投放Agent（降低买量人力成本）；这三个优势同时具备的公司屈指可数，掌阅是目前最接近"漫剧厂"而不是"内容方"的存在'),
  b2('对快手的含义：掌阅泡漫处于扩张期（短剧营收+139%），买量需求会快速提升；这是快手主动触达的最佳时机——掌阅有成熟的投放能力和清晰的ROI需求，不需要教育期，谈好CPM和转化数据就可以起量'),
  b2b('销售侧动作：', '主动约掌阅科技/泡漫团队，带着快手平台的漫剧用户画像数据（年龄/地域/使用时长）和竞对平台的价格对比；掌阅有自研投放Agent，他们自己会做归因，快手需要给的是"用户在哪里"而不是"怎么投"'),
  sp(),

  // D5 出海AI短剧
  h2('昆仑万维双产品FreeReels+DramaWave包揽出海TOP1/3，总素材量环比+4.33%'),
  b1('AI短剧出海的买量生态正在快速成熟，昆仑万维在建立自己的出海内容+分发壁垒；对快手国际化业务，这是一个值得研究的参照'),
  b2('FreeReels素材量18.1万组，DramaWave 14万组，两款产品目标市场分别是印尼/印度/巴西和欧洲，策略上是地域分层；这意味着AI短剧的出海不是一个单一市场的机会，是多地域差异化策略'),
  b2b('对快手国际化的含义：', 'Kwai的主要市场（巴西、印尼）和FreeReels的核心市场高度重合；如果昆仑万维正在Kwai平台买量，快手国际化团队应该已经在服务这个客户；如果没有，这是个明确的漏单——出海AI短剧厂商的买量需求和Kwai用户池有天然匹配'),

  new Paragraph({ children: [new TextRun('')], spacing: { before: 600 } }),
  new Paragraph({
    children: [new TextRun({ text: '内容消费研究组 · 数据仅供内部参考 · 2026 W17', size: 18, color: '94a3b8', font: '宋体' })],
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: 'e2e8f0' } },
    spacing: { before: 160 },
  }),
];

const doc = new Document({
  numbering,
  styles: {
    default: { document: { run: { font: '宋体', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: '宋体', color: '0f1f3d' },
        paragraph: { spacing: { before: 560, after: 180 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: '宋体', color: '1a2f52' },
        paragraph: { spacing: { before: 400, after: 100 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/Users/zhuyiqi/Documents/工作空间-测试1/game-intel-hub/W17行业深度研判.docx', buf);
  console.log('Done: W17行业深度研判.docx');
});
