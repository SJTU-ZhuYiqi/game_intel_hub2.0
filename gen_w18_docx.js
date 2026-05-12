const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

// ── 颜色常量 ──
const C = {
  darkBlue: "1F3864",
  midBlue: "2E75B6",
  lightBlue: "D5E8F0",
  lightBlue2: "E8F4FD",
  orange: "C55A11",
  lightOrange: "FCE4D6",
  red: "C00000",
  lightRed: "FFE0E0",
  violet: "7030A0",
  lightViolet: "EDE7F6",
  green: "375623",
  lightGreen: "E2EFDA",
  teal: "215868",
  lightTeal: "DAEEF3",
  gray1: "404040",
  gray2: "595959",
  gray3: "808080",
  grayBg: "F2F2F2",
  white: "FFFFFF",
  black: "000000",
};

// ── 通用样式 ──
const border = (color = "CCCCCC") => ({ style: BorderStyle.SINGLE, size: 4, color });
const noBorder = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
const allBorders = (color = "CCCCCC") => ({ top: border(color), bottom: border(color), left: border(color), right: border(color) });
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// 全宽 (A4 内容宽度，2cm margins each side: 11906 - 2*1134 = 9638 DXA ≈ 用9360)
const TW = 9360;

function hRun(text, opts = {}) {
  return new TextRun({ text, font: "Arial", ...opts });
}

function para(children, opts = {}) {
  if (typeof children === 'string') children = [hRun(children)];
  return new Paragraph({ children, spacing: { after: 80 }, ...opts });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", bold: true, size: 32 })],
    spacing: { before: 320, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.midBlue, space: 4 } },
  });
}

function h2(text, color = C.darkBlue) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", bold: true, size: 26, color })],
    spacing: { before: 240, after: 120 },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", bold: true, size: 22, color: C.gray1 })],
    spacing: { before: 160, after: 80 },
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    children: [hRun(text, { size: 20 })],
    spacing: { after: 60 },
  });
}

function bodyPara(text) {
  return new Paragraph({
    children: [hRun(text, { size: 20, color: C.gray1 })],
    spacing: { after: 100 },
  });
}

function boldLabel(label, text) {
  return new Paragraph({
    children: [
      hRun(label + '：', { bold: true, size: 20 }),
      hRun(text, { size: 20, color: C.gray1 }),
    ],
    spacing: { after: 80 },
  });
}

function callout(title, body, action, bgColor, borderColor) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [TW],
    rows: [
      new TableRow({
        children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 16, color: borderColor }, bottom: border("CCCCCC"), left: { style: BorderStyle.SINGLE, size: 16, color: borderColor }, right: border("CCCCCC") },
          width: { size: TW, type: WidthType.DXA },
          shading: { fill: bgColor, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [
            new Paragraph({ children: [hRun(title, { bold: true, size: 22, color: borderColor })], spacing: { after: 60 } }),
            new Paragraph({ children: [hRun(body, { size: 20, color: C.gray1 })], spacing: { after: 80 } }),
            new Paragraph({ children: [hRun('→ 建议动作：', { bold: true, size: 20, color: borderColor }), hRun(action, { size: 20, italic: true, color: C.gray1 })], spacing: { after: 0 } }),
          ],
        })],
      }),
    ],
  });
}

function sectionDivider() {
  return new Paragraph({ spacing: { before: 120, after: 120 } });
}

// 表格辅助
function makeTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders: allBorders(C.midBlue),
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: C.midBlue, type: ShadingType.CLEAR },
      margins: cellMargins,
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ children: [hRun(h, { bold: true, size: 18, color: C.white })], alignment: AlignmentType.CENTER })],
    })),
  });
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders: allBorders("CCCCCC"),
      width: { size: colWidths[ci], type: WidthType.DXA },
      shading: { fill: ri % 2 === 0 ? C.white : C.grayBg, type: ShadingType.CLEAR },
      margins: cellMargins,
      children: [new Paragraph({ children: [hRun(cell, { size: 18 })], alignment: AlignmentType.LEFT })],
    })),
  }));
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// 封面
function coverSection() {
  return [
    new Paragraph({ spacing: { before: 2000 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('内容消费行业', { size: 52, bold: true, color: C.midBlue })],
      spacing: { after: 60 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('深度研判报告', { size: 64, bold: true, color: C.darkBlue })],
      spacing: { after: 120 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('W18 · 2026年5月4日—5月10日', { size: 28, color: C.gray2 })],
      spacing: { after: 60 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('游戏行业 × 短剧/漫剧行业', { size: 22, color: C.gray3 })],
      spacing: { after: 0 },
    }),
    new Paragraph({ spacing: { before: 60, after: 60 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.midBlue, space: 1 } } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('发布日期：2026年5月11日', { size: 20, color: C.gray3 })],
      spacing: { after: 40 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('商业化中心 · 数据分析', { size: 20, color: C.gray3 })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// 执行摘要
function execSummary() {
  return [
    h1('执行摘要'),
    bodyPara('本周（2026年W18，5月4日—5月10日）游戏+短剧/漫剧行业呈现出以下核心判断：'),
    sectionDivider(),

    // 游戏摘要框
    new Table({
      width: { size: TW, type: WidthType.DXA },
      columnWidths: [TW],
      rows: [new TableRow({ children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 20, color: C.midBlue }, bottom: border("CCCCCC"), left: border("CCCCCC"), right: border("CCCCCC") },
        width: { size: TW, type: WidthType.DXA },
        shading: { fill: C.lightBlue2, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ children: [hRun('🎮  游戏行业', { bold: true, size: 26, color: C.midBlue })], spacing: { after: 100 } }),
          new Paragraph({ children: [hRun('核心结论：', { bold: true, size: 20 }), hRun('异环全球首日流水破亿，但PC+主机贡献超75%，"大屏化"趋势实锤；小游戏掌上谈兵稳守前三，下沉市场持续分流；4月版号创近期新高，叠加暑假备战，Q3买量高峰可期。', { size: 20, color: C.gray1 })], spacing: { after: 80 } }),
          new Paragraph({ children: [hRun('关键数字：', { bold: true, size: 20 }), hRun('首日破亿 | 147款版号 | 日均买量破万条 | 莉莉丝70+起反腐案', { size: 20, color: C.gray1 })], spacing: { after: 0 } }),
        ],
      })]})]
    }),
    sectionDivider(),

    // 短剧摘要框
    new Table({
      width: { size: TW, type: WidthType.DXA },
      columnWidths: [TW],
      rows: [new TableRow({ children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 20, color: C.orange }, bottom: border("CCCCCC"), left: border("CCCCCC"), right: border("CCCCCC") },
        width: { size: TW, type: WidthType.DXA },
        shading: { fill: C.lightOrange, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ children: [hRun('📺  短剧 / 漫剧行业', { bold: true, size: 26, color: C.orange })], spacing: { after: 100 } }),
          new Paragraph({ children: [hRun('核心结论：', { bold: true, size: 20 }), hRun('听花岛×掌玩10亿押注海外AI短剧（FlickReels已验证美区#1）；AI微短剧产量占比超95%，制作端AI化完成；抖音漫剧分成大幅下调，部分创作者将外流；最高检明确AI短剧盗录入罪，版权框架向AI内容延伸。', { size: 20, color: C.gray1 })], spacing: { after: 80 } }),
          new Paragraph({ children: [hRun('关键数字：', { bold: true, size: 20 }), hRun('10亿出海投入 | 红果月活2563万 | 仿真人剧分成60→40 | 五一档6.6亿', { size: 20, color: C.gray1 })], spacing: { after: 0 } }),
        ],
      })]})]
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ── Part 1: 游戏行业 ──
function gamePart() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('第一部分', { size: 20, color: C.gray3 })],
      spacing: { before: 160, after: 40 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('游戏行业深度研判', { bold: true, size: 40, color: C.midBlue })],
      spacing: { after: 40 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('2026年W18（5月4日—5月10日）', { size: 22, color: C.gray2 })],
      spacing: { after: 0 },
    }),
    new Paragraph({ spacing: { before: 60, after: 60 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.midBlue } } }),
    sectionDivider(),

    // 1. 核心KPI
    h1('一、本周核心数据'),
    h2('1.1 榜单格局（截至5月2日）', C.midBlue),
    makeTable(
      ['榜单', 'Top 1', 'Top 2', 'Top 3', '关键动态'],
      [
        ['iOS畅销榜', '王者荣耀', '和平精英', '三角洲行动', '腾讯三产品包揽前三，格局稳固'],
        ['微信小游戏', '掌上谈兵', '疯狂水世界', '向僵尸开炮', '五一后格局稳定，下沉市场固化'],
      ],
      [1600, 1600, 1600, 1600, 2960]
    ),
    sectionDivider(),

    h2('1.2 4月全球手游收入TOP榜（AppMagic）', C.midBlue),
    makeTable(
      ['全球排名', '产品', '4月收入（美元）', '环比变化', '快手相关性'],
      [
        ['#1', '王者荣耀', '~1.38亿', '重回榜首', '腾讯核心，关注版本活动节点'],
        ['#2', 'Last War: Survival', '>1亿', '从峰值回落', '世纪华通旗下，快手有稳定份额'],
        ['#3', '无尽冬日', '>1亿', '稳定', '世纪华通，快手买量主力产品之一'],
        ['#11-15', '三角洲行动', '~5100万', '上升4位', '腾讯FPS，内容营销窗口'],
        ['#11-15', '崩坏:星穹铁道', '~3680万', '较1月低谷大幅回升', '米哈游，二游买量重启信号'],
      ],
      [1200, 2000, 1600, 1600, 2960]
    ),
    sectionDivider(),

    // 2. 异环深度复盘
    h1('二、重磅事件深度复盘：《异环》全球首日'),
    h2('2.1 核心数据拆解', C.midBlue),
    makeTable(
      ['维度', '数据', '分析含义'],
      [
        ['首日全平台流水', '破1亿元人民币', '完美世界年内最大验证，超市场预期'],
        ['国内PC端流水占比', '~65%', 'iOS畅销榜（约第61位）严重低估真实规模'],
        ['海外PC+PS主机占比', '~75%', '"大屏化"趋势首次被国产二游数据实锤'],
        ['海外覆盖范围', '170+国家/地区', '日本、美国为流水主力；韩国超预期'],
        ['iOS免费榜首日', '148个区域TOP50', '开测第2日，用户渗透广但付费转化偏弱'],
      ],
      [2500, 2500, 4360]
    ),
    sectionDivider(),

    h2('2.2 分析师研判', C.midBlue),
    callout(
      '判断一：iOS畅销榜≠真实规模 — "大屏化"趋势已实锤',
      '过去业界习惯用iOS畅销榜衡量游戏流水，异环用首日数据证明这套标准对"大屏化"产品失效。当PC端贡献65%国内流水、PC+主机贡献75%海外流水时，移动端数据只是全貌的冰山一角。这一结构性变化对买量策略意义重大：移动端买量ROI将持续低于大屏端投放，但移动端仍是社区传播和种草的核心渠道。',
      '异环在快手的核心价值在内容营销（直播/短视频种草）而非重度买量采买。',
      C.lightViolet, C.violet
    ),
    sectionDivider(),
    callout(
      '判断二：异环验证"都市题材"可行性 — 下一个赛道窗口',
      '异环是近年极少数成功跑通"现代都市+开放世界"题材的国产二游。玩家对都市模拟生活玩法（开车、租房、穿搭）反馈超预期，付费重心计划从标准二游卡池迁移至道具/皮肤/联名，理论上可覆盖更宽的付费用户群。若后续版本迭代稳定，异环可能成为完美世界扭转基本面的关键变量。',
      '持续关注异环月留存和付费深度数据，评估Q3买量预算规模是否扩大。',
      C.lightBlue2, C.midBlue
    ),
    sectionDivider(),

    h2('2.3 后续关键节点', C.midBlue),
    bullet('5月7日：首个新角色「浔」（S级限定）卡池上线，首个付费转化关键检验点'),
    bullet('5月8日：保时捷跑车联名活动（都市题材特有联名，覆盖车主/时尚用户群）'),
    bullet('后续版本：人气角色「安魂曲」上线 + 地图扩充 + 长线付费体系构建'),
    sectionDivider(),

    // 3. 小游戏
    h1('三、小游戏市场：下沉格局固化'),
    h2('3.1 五一后微信小游戏TOP5格局', C.midBlue),
    makeTable(
      ['排名', '产品', '厂商', '类型', '买量状态', '快手机会评级'],
      [
        ['#1', '掌上谈兵', '聚力网络', '三国卡牌RPG', '日均破万条（持续）', '★★★★★'],
        ['#2', '疯狂水世界', '益世界', '经营+SLG', '双端活跃', '★★★★☆'],
        ['#3', '向僵尸开炮', '盛昌网络', '益智塔防', '长线IAA稳定', '★★★☆☆'],
        ['#4', '无尽冬日', '点点互动', 'SLG对战', '长线重投放', '★★★★☆'],
        ['#5', '三国：冰河时代', '欢游互动', 'SLG', '五一后新晋', '★★★☆☆'],
      ],
      [800, 1800, 1500, 1500, 1800, 1960]
    ),
    sectionDivider(),

    h2('3.2 掌上谈兵：本周最高优先级买量目标', C.midBlue),
    callout(
      '掌上谈兵 — 立即行动',
      '4月24日-5月1日持续霸榜微信小游戏畅销榜#1，之后稳守#3，与无尽冬日、三国冰河时代共同构成五一后三强。买量侧：4月18日起日均投放破万条，预算充沛且采买效率稳定（否则不可能维持超2周的高强度投放）。目标用户30-45岁三四线城市男性，与快手主力用户高度重合。这是本周可立即触达的最高优先级广告主。',
      '推进IAP扩量方案，评估快手平台流量包/品牌合作可能性。',
      C.lightGreen, C.green
    ),
    sectionDivider(),

    // 4. 行业动态
    h1('四、行业重大动态'),
    h2('4.1 版号政策：4月147款创近期新高', C.midBlue),
    bodyPara('4月29日，国家新闻出版署公布4月审批信息：147款国产游戏+7款进口游戏版号过审，较3月133款增加10.5%，整体维持高位。'),
    makeTable(
      ['重点产品', '厂商', '类型', 'TapTap预约/评分', '上线预判'],
      [
        ['遗忘之海', '网易', '开放世界RPG', '83万预约 / 8.8分', '6月前 — 最值得关注的开放世界新品'],
        ['洛克王国：精灵牌', '腾讯', '卡牌', '—', '腾讯IP矩阵拓展'],
        ['冰雪之笼', '巨人网络', '—', '—', '巨人Q2重点产品'],
        ['绮梦异谭', '友谊时光', '女性向', '—', '女性向小众市场'],
        ['弧光猎人（进口）', '腾讯', '射击', '—', '腾讯代理进口射击产品'],
      ],
      [2000, 1500, 1500, 2000, 2360]
    ),
    sectionDivider(),

    h2('4.2 韩国OneStore×腾讯：小游戏出海新渠道', C.midBlue),
    bodyPara('4月30日，韩国OneStore（SKT/KT/LG U+三大运营商+Naver联合体）宣布与腾讯深度合作，5月起通过"One Play Game"服务引进中国微信小游戏。'),
    makeTable(
      ['维度', '数据', '对快手的含义'],
      [
        ['OneStore市场份额', '韩国约15%（Google Play约70%）', '—'],
        ['用户ARPPU', 'Google Play的约5倍', '韩国小游戏用户付费能力显著高于预期'],
        ['引进规模', '5月起大规模引进', '国内掌上谈兵、无尽冬日等产品可能优先出海韩国'],
        ['与快手相关性', '出海渠道验证', '快手海外版（Kwai）在东南亚可关注类似机会'],
      ],
      [2000, 2400, 4960]
    ),
    sectionDivider(),

    h2('4.3 莉莉丝反腐通报：行业合规管理进入新阶段', C.midBlue),
    bodyPara('5月6日，莉莉丝游戏官方账号发布《廉洁自律通报》，是近年国内头部游戏厂商中力度最大的反腐通报。'),
    bullet('查处规模：70+起案件，涉及采购/运营/发行/质量管理等核心岗位'),
    bullet('人员处置：22人辞退，10+人移送公安机关（含发行负责人、质管负责人等中高层）'),
    bullet('供应商黑名单：首次公开22家涉及商业贿赂/利益输送的供应商名单'),
    bodyPara(''),
    callout(
      '研判：莉莉丝反腐是行业信号，供应商/服务商管理将进入强监管期',
      '莉莉丝此举的行业示范效应显著。其他头部厂商面临"跟进压力"（否则引发舆论质疑）。短期影响：存量供应商合规审查周期拉长，新合作洽谈流程趋严。中期影响：游戏行业整体供应链管理将向"白名单制"演进，劣质流量/数据注水服务商首当其冲。对快手而言，这是加速推进"正规合规"游戏合作方资质认证的有利窗口。',
      '在游戏合作方沟通中主动强调快手平台的合规生态，建立差异化信任优势。',
      C.lightRed, C.red
    ),
    sectionDivider(),

    h2('4.4 AI游戏赛道加速', C.midBlue),
    makeTable(
      ['事件', '核心信息', '影响判断'],
      [
        ['Astrocade融资5600万美元', '李飞飞联创，红杉领投，谷歌/英伟达/LG跟投；8个月2000万用户', '"零代码生成游戏"大资本入场，AI游戏赛道确认'],
        ['恺英×中手游GamePartner.AI', '5月8日发布，AI辅助游戏开发+发行一站式工具', '国内AI游戏工具赛道正式起量'],
        ['Supercell全资收购Metacore', '整合Merge Mansion，腾讯系休闲赛道补全', '腾讯全球休闲游戏矩阵进一步完善'],
        ['网易《归唐》亮相', '雷火工作室，线性叙事，B站1564万播放', '网易下半年大作储备，非买量敏感型产品'],
      ],
      [2000, 3500, 3860]
    ),
    sectionDivider(),

    // 5. 游戏新游预告
    h1('五、新游预告与买量节点'),
    makeTable(
      ['上线日期', '产品', '厂商', '类型', 'TapTap评分', '优先级'],
      [
        ['5月15日', '镭明闪击', '朝夕光年', '竞技射击', '—', '中 — 观察首周数据'],
        ['5月20日', '宗师之上', '广州灵犀', '放置RPG', '9.1', '高 — 预约热度榜第6，用户期待强'],
        ['6月前后', '遗忘之海', '网易', '开放世界RPG', '8.8', '极高 — TapTap预约83万，年内最值期待'],
        ['待定', '无限大', '网易', '开放世界二游', '—', '中高 — 竞品标杆，对异环形成压制'],
      ],
      [1500, 1800, 1800, 1800, 1500, 1960]
    ),
    sectionDivider(),

    // 6. 游戏研判总结
    h1('六、核心洞察与建议行动'),
    sectionDivider(),
    callout(
      '洞察1（立即行动）：异环大屏化验证 — iOS排名≠真实规模，内容营销优于买量',
      '异环首日破亿但iOS约第61位，PC+主机贡献75%以上流水。对快手而言，异环的直播/短视频内容营销价值显著高于移动端买量ROI。重点布局游戏主播合作和用户UGC内容放量，而非采买CPC/CPA。',
      '评估快手游戏直播专区与异环的内容营销合作方案，测算内容投入与用户增量ROI。',
      C.lightViolet, C.violet
    ),
    sectionDivider(),
    callout(
      '洞察2（立即行动）：掌上谈兵是本周最高优先级触达目标',
      '五一后稳守前三+日均买量破万条，目标用户与快手高度重合。这款三国卡牌RPG的广告主预算充沛，是可立即推进IAP扩量方案的最优先广告主。',
      '本周优先安排掌上谈兵扩量评估，推进快手专属流量包方案。',
      C.lightGreen, C.green
    ),
    sectionDivider(),
    callout(
      '洞察3（版号信号）：4月版号147款创高，Q3买量高峰将至',
      '版号放量+新游备战+暑假效应三重叠加，Q2末-Q3初将出现新游密集上线买量高峰。网易《遗忘之海》（预约83万）是最高关注度产品，《宗师之上》《镭明闪击》是近期上线产品。',
      '提前锁定网易《遗忘之海》上线节点的流量资源，确保Q3买量高峰期的快手份额。',
      C.lightBlue2, C.midBlue
    ),
    sectionDivider(),
    callout(
      '洞察4（行业趋势）：AI游戏赛道大资本入场 — 平台AI分发能力建设窗口',
      'Astrocade（"游戏界的抖音"）5600万美元B轮+国内GamePartner.AI同步起量，AI生成游戏赛道正式进入规模化。这类产品的用户生成内容天然适合短视频分发，是快手游戏内容生态的潜在新增量。',
      '关注AI游戏工具赛道进展，评估平台对AI生成游戏内容的分发能力和商业化策略。',
      C.lightOrange, C.orange
    ),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ── Part 2: 短剧/漫剧 ──
function dramaPart() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('第二部分', { size: 20, color: C.gray3 })],
      spacing: { before: 160, after: 40 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('短剧 / 漫剧行业深度研判', { bold: true, size: 40, color: C.orange })],
      spacing: { after: 40 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [hRun('2026年W18（5月4日—5月10日）', { size: 22, color: C.gray2 })],
      spacing: { after: 0 },
    }),
    new Paragraph({ spacing: { before: 60, after: 60 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } } }),
    sectionDivider(),

    // 1. 平台数据
    h1('一、平台数据总览'),
    h2('1.1 漫剧平台月活（截至2026年3月）', C.orange),
    makeTable(
      ['平台', '月活用户', '用户结构特征', '关键动态'],
      [
        ['红果免费漫剧', '2563.3万', '男性/90后/中老年为主', '稳步增长，已成漫剧第一平台'],
        ['火龙漫剧', '500万+', '男性为主（上线次月）', '快速起量，新入场玩家'],
      ],
      [2200, 1600, 2800, 2760]
    ),
    sectionDivider(),

    h2('1.2 W18网播剧热榜（4月25日—5月1日）', C.orange),
    makeTable(
      ['排名', '剧集', '播放平台', '播映指数', '出品方', '环比变化'],
      [
        ['#1', '黑夜告白', '优酷', '79', '上海儒意/优酷/万达', '新晋'],
        ['#2', '蜜语纪', '爱奇艺/腾讯', '78', '腾讯/爱奇艺', '↓1'],
        ['#3', '佳偶天成', '爱奇艺/腾讯', '76.9', '爱奇艺/欢瑞世纪', '新晋'],
        ['#4', '白日提灯', '腾讯视频', '74.2', '腾讯/万众合星', '—'],
        ['#5', '逐玉', '腾讯/爱奇艺', '74.1', '爱奇艺/浩瀚影视', '↓3'],
      ],
      [800, 1600, 1600, 1200, 2200, 1960]
    ),
    sectionDivider(),

    h2('1.3 W18综艺热榜（4月25日—5月1日）', C.orange),
    makeTable(
      ['排名', '综艺', '播映平台', '播映指数', '出品方'],
      [
        ['#1', '乘风2026', '芒果TV', '78.4', '芒果TV'],
        ['#2', '怦然心动20岁：冬季', '优酷', '70.1', '优酷'],
        ['#3', '超燃青春的合唱', '爱奇艺', '66.1', '爱奇艺'],
        ['#4', '魔力歌先生', '腾讯视频', '65.3', '腾讯视频'],
      ],
      [800, 2500, 1600, 1500, 2960]
    ),
    sectionDivider(),

    // 2. 重磅事件
    h1('二、重磅事件深度复盘'),
    h2('2.1 听花岛×掌玩：10亿押注海外AI短剧', C.orange),
    bodyPara('2026年4月27日，听花岛（国内短剧头部制作方）与掌玩（短剧发行商）联合宣布启动"打造下一代Netflix"海外合作计划，计划投入10亿元人民币布局海外AI短剧业务。'),
    makeTable(
      ['维度', '详情'],
      [
        ['合作主体', '听花岛（制作）+ 掌玩（发行）'],
        ['资金规模', '10亿元人民币'],
        ['已验证成果', 'FlickReels平台曾登美国AppStore娱乐榜#1、全榜第二'],
        ['资金用途', '剧本开发 + 成品剧制作两大方向（各约50%）'],
        ['核心优势', '听花岛爆款内容储备 + 掌玩海外发行渠道（主打北美/东南亚）'],
      ],
      [2000, 7360]
    ),
    sectionDivider(),
    callout(
      '深度研判：10亿资金背后是"出海模式已验证"的强信心',
      '听花岛×掌玩的10亿投入不是探索性资金，而是已有FlickReels登顶美区的成功经验背书。这意味着：①中国AI短剧在海外市场的商业模式已跑通（内容+买量驱动）；②头部玩家已在规模化阶段——下一场竞争是效率之争，而非模式探索。对快手而言：Kwai（快手海外版）在东南亚/拉美的用户基础，与AI短剧的目标增量市场存在结构性匹配。但核心难点是内容本地化能力，而非流量本身。',
      '评估快手海外平台与听花岛/掌玩等国内AI短剧头部厂商的流量合作可能性，探索差异化出海分发通路。',
      C.lightOrange, C.orange
    ),
    sectionDivider(),

    h2('2.2 抖音漫剧分成系数大幅下调（4月起生效）', C.orange),
    bodyPara('抖音集团短剧版权中心宣布，自2026年4月起正式调整漫剧内容分成系数：'),
    makeTable(
      ['内容类型', '原分成系数', '新分成系数', '下调幅度', '分析含义'],
      [
        ['仿真人剧', '60', '40', '▼33%', '主流AI生成内容，受影响最大'],
        ['3D动画漫剧', '50', '40', '▼20%', '制作成本较高，利润空间进一步压缩'],
        ['动态解说漫剧', '5', '1', '▼80%', '几乎等于断档，将快速消亡'],
        ['2D动画漫剧', '40', '40', '持平', '传统漫画改编，暂保护'],
        ['表情包动态漫剧', '10', '10', '持平', '低成本内容，维持基本激励'],
      ],
      [1800, 1600, 1600, 1400, 2960]
    ),
    sectionDivider(),
    callout(
      '深度研判：平台从"供给激励"转向"平台获利"，内容创作者面临利润压缩',
      '分成系数下调是明确的信号：抖音认为漫剧供给已充足，不再需要高分成激励创作者；平台正在从"培育期分成"转向"成熟期收割"逻辑。对创作者端：仿真人剧是目前最主流的AI生成漫剧类型，33%分成下调将直接压缩毛利率（原本AI生产成本已很低，但买量成本高）。对快手而言：这是吸引被抖音分成下调驱逐的漫剧创作者入驻快手平台的窗口期。如果快手能维持或提高漫剧分成系数，可能在3-6个月内获得一批迁移创作者。',
      '研究快手当前漫剧分成政策与抖音差距，评估针对性扶持政策吸引被驱逐创作者的可行性与成本。',
      C.lightRed, C.red
    ),
    sectionDivider(),

    h2('2.3 AI微短剧产量占比超95%：制作端AI化完成', C.orange),
    bodyPara('上海市政府信息披露：当前AI微短剧产量占比已超95%。这意味着制作端的AI化进程实际上已经完成。'),
    callout(
      '深度研判：制作端效率竞争到顶，版权与分发成为新核心变量',
      'AI产量超95%意味着：任何人都可以以极低成本生产漫剧内容，"制作能力"不再是竞争壁垒。下一阶段的竞争将在三个维度展开：①优质IP版权——知名漫画/小说改编仍有内容护城河，AI难以复制；②用户分发效率——谁的平台能以最低eCPM触达最精准的漫剧消费用户；③内容质量差异化——在AI均质化内容海洋中，精品仍有用户溢价。快手的核心优势在②（下沉用户精准匹配），但①和③需要内容合作生态支撑。',
      '重点评估快手短剧分发侧核心指标（完播率/分享率/付费率）与行业基准对比，建立分发效率竞争壁垒。',
      C.lightViolet, C.violet
    ),
    sectionDivider(),

    h2('2.4 最高检：AI短剧盗录牟利构成犯罪（5月9日）', C.orange),
    bodyPara('2026年5月9日，最高人民检察院官方渠道发文：AI短剧盗录牟利行为构成侵犯著作权罪。这是官方首次明确将AI生成内容的盗录行为纳入刑事规制框架。'),
    bullet('法律含义：AI内容的著作权保护边界正式向数字版权领域延伸，AI短剧版权归属问题有了明确司法立场'),
    bullet('行业影响：短期内将压制盗录搬运内容（打击"搬运号"、"镜像站"等灰色渠道），正版内容分发平台受益'),
    bullet('对快手的影响：若快手平台存在盗录AI短剧内容的创作者，需及时清查；正版授权内容的投入产出比将相应提高'),
    sectionDivider(),

    // 3. 院线市场
    h1('三、院线市场：五一档6.6亿元'),
    makeTable(
      ['指标', '数据', '同比', '含义'],
      [
        ['五一档总票房（含预售）', '6.6亿元', '—', '五一档规模正常，但集中度高（两片占6成）'],
        ['观影人次', '1814.1万', '+6.4%', '人次增长，但平均票价下滑'],
        ['首日平均票价', '36.9元', '—', '创近四年五一档最低，价格战已至底部'],
        ['档期领跑影片', '《消失的人》约2.25亿', '—', '悬疑类型爆款，首日后口碑驱动逆跌'],
        ['第二名', '《寒战1994》约1.77亿', '—', '港式警匪续集，IP效应稳定'],
      ],
      [2200, 1800, 1200, 4160]
    ),
    sectionDivider(),

    // 4. 核心洞察
    h1('四、核心洞察与建议行动'),
    sectionDivider(),
    callout(
      '洞察1（出海机会）：10亿押注验证海外AI短剧窗口期已开',
      '听花岛×掌玩的10亿投入+FlickReels美区#1验证，标志中国AI短剧出海进入规模化阶段。快手海外版（Kwai）在东南亚/拉美拥有规模用户基础，与AI短剧的增量市场高度匹配。核心壁垒不是流量，而是内容本地化能力和版权合规体系。',
      '评估与听花岛/FlickReels等国内头部海外短剧平台的流量合作方案，探索差异化分发渠道建立先发优势。',
      C.lightOrange, C.orange
    ),
    sectionDivider(),
    callout(
      '洞察2（创作者争夺）：抖音漫剧分成下调创造迁移窗口',
      '仿真人剧33%、动态解说80%的分成下调，将在3-6个月内逐步压缩创作者利润，部分创作者将主动寻求其他平台。快手如能在此窗口期提供更优厚的分成或流量扶持，有机会以较低成本获取优质漫剧创作者资源。',
      '研究快手与抖音漫剧分成差距，设计针对性迁移激励方案，优先吸引仿真人剧和3D动画创作者。',
      C.lightRed, C.red
    ),
    sectionDivider(),
    callout(
      '洞察3（AI渗透）：AI产量超95%，分发效率是快手的核心竞争窗口',
      '制作端壁垒消失后，哪个平台能以最低成本触达最精准的漫剧消费用户，才是真正的竞争优势。快手下沉用户与漫剧/短剧核心受众（男性/90后/中老年）高度重合，这是天然的分发效率优势——但需要量化数据支撑，并转化为面向广告主/创作者的差异化主张。',
      '量化快手漫剧内容的完播率/分享率/付费转化率与行业基准对比，建立"快手用户=漫剧精准受众"的数据叙事。',
      C.lightViolet, C.violet
    ),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ── 附录 ──
function appendix() {
  return [
    h1('附录：数据来源说明'),
    makeTable(
      ['数据类别', '来源', '覆盖周期'],
      [
        ['游戏全球收入榜', 'AppMagic（月度数据）', '2026年4月'],
        ['iOS畅销/免费榜', '七麦数据', '2026年4月26日—5月2日'],
        ['微信小游戏榜', '微信平台官方榜单', '五一档期'],
        ['版号过审信息', '国家新闻出版署', '2026年4月'],
        ['网播剧/综艺指数', '艺恩数据（播映指数）', '2026年4月25日—5月1日'],
        ['漫剧平台月活', 'QuestMobile 2026中国移动互联网春季大报告', '截至2026年3月'],
        ['五一档票房数据', '猫眼专业版', '截至2026年5月4日20:30'],
        ['行业新闻', '开源证券传媒周报（2026.05.05）、短剧自习室、DataEye研究院', '2026年4月27日—5月10日'],
        ['政策信息', '最高人民检察院官方、上海市政府信息公开', '2026年5月'],
      ],
      [2500, 4000, 2860]
    ),
    sectionDivider(),
    bodyPara('免责声明：本报告内容仅供内部参考，数据来源于公开信息，不构成投资建议。'),
  ];
}

// ── 构建文档 ──
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: C.darkBlue },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: C.midBlue },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: C.gray1 },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }, // 2cm
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [
            hRun('内容消费行业深度研判报告  W18 · 2026年5月4日—5月10日', { size: 16, color: C.gray3 }),
          ],
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
          spacing: { after: 0 },
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            hRun('第 ', { size: 16, color: C.gray3 }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: C.gray3 }),
            hRun(' 页  |  商业化中心 · 数据分析  |  仅供内部参考', { size: 16, color: C.gray3 }),
          ],
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
          spacing: { before: 0 },
        })],
      }),
    },
    children: [
      ...coverSection(),
      ...execSummary(),
      ...gamePart(),
      ...dramaPart(),
      ...appendix(),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('W18行业深度研判.docx', buf);
  console.log('✅ W18行业深度研判.docx 已生成');
});
