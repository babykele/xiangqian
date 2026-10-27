/* 云宝纹样设计工坊：可测试的创意解释与 SVG 图元。 */
(function (root) {
  'use strict';

  const COLOR_WORDS = [
    [/深蓝|藏蓝/, '#174B8B', '深蓝'],
    [/蓝|azure/i, '#2878D0', '云蓝'],
    [/黄|金/, '#F4B942', '暖黄'],
    [/红/, '#E45757', '珊瑚红'],
    [/粉/, '#EE7FA7', '樱花粉'],
    [/绿/, '#39A878', '青绿'],
    [/紫/, '#7B61C9', '紫罗兰'],
    [/橙/, '#ED8A3A', '橙色'],
    [/白/, '#F8FAFF', '云白'],
    [/黑|灰/, '#344054', '墨灰']
  ];

  const MOTIFS = [
    [/云|云朵/, 'cloud', '云朵'],
    [/星|星星|星光/, 'star', '星星'],
    [/叶|叶子|树叶/, 'leaf', '叶片'],
    [/爱心|心形|红心/, 'heart', '爱心'],
    [/菱形|方块|几何/, 'diamond', '菱形'],
    [/水滴|雨滴|浪花|波浪/, 'drop', '水滴']
  ];

  const DEFAULT = {
    motif: 'cloud', motifName: '云朵', accent: 'star', accentName: '星星',
    primary: '#2878D0', primaryName: '云蓝', secondary: '#F4B942', secondaryName: '暖黄',
    motion: 'translate', motionName: '平移', step: 78, direction: 'right',
    mood: '活泼', assumptions: [], recognized: []
  };

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function firstMatch(text, list) {
    for (const item of list) if (item[0].test(text)) return item;
    return null;
  }

  function orderedMatches(text, list) {
    return list.map(item => {
      const match = text.match(item[0]);
      return match ? { item, index: match.index, text: match[0], end: match.index + match[0].length } : null;
    }).filter(Boolean).sort((a, b) => a.index - b.index);
  }

  function pairedColor(text, motifMatch, colorMatches, used) {
    const candidates = colorMatches.filter((color, index) => !used.has(index) && color.end <= motifMatch.index && motifMatch.index - color.end <= 2);
    if (!candidates.length) return null;
    const chosen = candidates[candidates.length - 1];
    const index = colorMatches.indexOf(chosen);
    used.add(index);
    return chosen.item;
  }

  function interpretIdea(input) {
    const text = String(input || '').trim();
    const result = JSON.parse(JSON.stringify(DEFAULT));
    result.assumptions = [];
    result.recognized = [];

    const motifEntries = orderedMatches(text, MOTIFS);
    const motifMatches = motifEntries.map(entry => entry.item);
    if (motifMatches[0]) {
      result.motif = motifMatches[0][1]; result.motifName = motifMatches[0][2];
      result.recognized.push('单位图案：' + result.motifName);
    } else result.assumptions.push('未识别具体图案，暂用云朵作为单位图案');
    if (motifMatches[1]) {
      result.accent = motifMatches[1][1]; result.accentName = motifMatches[1][2];
      result.recognized.push('搭配图案：' + result.accentName);
    } else if (result.motif !== 'star') {
      result.accent = 'star'; result.accentName = '星星';
      result.assumptions.push('未说明搭配图案，暂配星星');
    } else {
      result.accent = 'cloud'; result.accentName = '云朵';
    }

    const colorEntries = orderedMatches(text, COLOR_WORDS);
    const usedColors = new Set();
    const primaryPair = motifEntries[0] ? pairedColor(text, motifEntries[0], colorEntries, usedColors) : null;
    const secondaryPair = motifEntries[1] ? pairedColor(text, motifEntries[1], colorEntries, usedColors) : null;
    const remainingColors = colorEntries.filter((entry, index) => !usedColors.has(index)).map(entry => entry.item);
    const primaryColor = primaryPair || remainingColors.shift();
    const secondaryColor = secondaryPair || remainingColors.shift();
    if (primaryColor) {
      result.primary = primaryColor[1]; result.primaryName = primaryColor[2];
      result.recognized.push(result.motifName + '颜色：' + result.primaryName);
    } else result.assumptions.push('未说明主色，暂用云蓝');
    if (secondaryColor) {
      result.secondary = secondaryColor[1]; result.secondaryName = secondaryColor[2];
      result.recognized.push(result.accentName + '颜色：' + result.secondaryName);
    } else if (primaryColor && primaryColor[1] !== '#F4B942') {
      result.assumptions.push('未说明辅色，暂用暖黄');
    }

    if (/旋转|转半圈|180/.test(text)) {
      result.motion = 'rotate'; result.motionName = '旋转180°后平移';
      result.recognized.push('变换方式：' + result.motionName);
    } else if (/翻转|翻折|轴对称|镜像|一正一反|交替/.test(text)) {
      result.motion = 'reflect'; result.motionName = '轴对称后平移';
      result.recognized.push('变换方式：' + result.motionName);
    } else if (/平移|重复|连续|一直排|排列/.test(text)) {
      result.motion = 'translate'; result.motionName = '平移';
      result.recognized.push('变换方式：平移');
    } else result.assumptions.push('未说明变换方式，暂按向右平移');

    if (/向左|往左|从右到左/.test(text)) result.direction = 'left';
    else if (/两边|两侧|左右|双向/.test(text)) result.direction = 'both';
    if (/(向左|往左|从右到左|向右|往右|从左到右|两边|两侧|左右|双向|水平方向)/.test(text)) {
      result.recognized.push('生成方向：' + (result.direction === 'left' ? '向左' : result.direction === 'both' ? '向左右两侧' : '向右'));
    } else result.assumptions.push('未说明生成方向，暂按向右排列');
    if (/紧密|挨着|没有空隙|无缝/.test(text)) result.step = 64;
    else if (/宽松|空一些|间隔大|留白/.test(text)) result.step = 98;
    else {
      const distance = text.match(/(?:间距|平移|每次).{0,5}(\d{2,3})/);
      if (distance) result.step = Math.max(56, Math.min(110, Number(distance[1])));
    }

    const mood = firstMatch(text, [
      [/活泼|可爱|童趣/, '', '活泼'], [/宁静|温柔|清新/, '', '清新'],
      [/科技|未来|酷/, '', '未来感'], [/热情|活力|阳光/, '', '有活力']
    ]);
    if (mood) { result.mood = mood[2]; result.recognized.push('氛围：' + mood[2]); }
    result.sourceText = text;
    result.confidence = Math.max(0.35, Math.min(0.96, 0.48 + result.recognized.length * 0.09 - result.assumptions.length * 0.04));
    return result;
  }

  function motifShape(kind, color, accent) {
    const fill = esc(color), line = esc(accent || '#17324D');
    if (kind === 'star') return `<path d="M0 -18 L5 -6 L18 -6 L8 2 L12 16 L0 8 L-12 16 L-8 2 L-18 -6 L-5 -6 Z" fill="${fill}" stroke="${line}" stroke-width="2"/>`;
    if (kind === 'leaf') return `<path d="M-20 10 Q-10 -20 20 -12 Q12 16 -20 10Z" fill="${fill}" stroke="${line}" stroke-width="2"/><path d="M-14 7 L13 -9" stroke="${line}" stroke-width="2"/>`;
    if (kind === 'heart') return `<path d="M0 17 C-25 2 -22 -17 -8 -18 C0 -18 3 -11 3 -9 C7 -18 15 -20 22 -15 C35 -5 21 10 0 17Z" fill="${fill}" stroke="${line}" stroke-width="2"/>`;
    if (kind === 'diamond') return `<path d="M0 -20 L22 0 L0 20 L-22 0Z" fill="${fill}" stroke="${line}" stroke-width="2"/><path d="M0 -11 L12 0 L0 11 L-12 0Z" fill="none" stroke="${line}" stroke-width="2"/>`;
    if (kind === 'drop') return `<path d="M0 -22 C10 -7 19 1 17 11 C15 23 -15 23 -17 11 C-19 1 -10 -7 0 -22Z" fill="${fill}" stroke="${line}" stroke-width="2"/>`;
    return `<path d="M-22 10 C-29 0 -20 -10 -11 -9 C-7 -21 12 -23 18 -10 C31 -11 35 6 24 12 L-17 12 C-19 12 -21 11 -22 10Z" fill="${fill}" stroke="${line}" stroke-width="2"/>`;
  }

  function motionTransform(motion, index) {
    if (index % 2 === 0) return '';
    if (motion === 'reflect') return 'scale(-1 1)';
    if (motion === 'rotate') return 'rotate(180)';
    return '';
  }

  function ruleSentence(config) {
    const direction = config.direction === 'left' ? '沿水平方向向左生成' : config.direction === 'both' ? '沿左右方向连续排列' : '沿水平方向向右生成';
    const operation = config.motion === 'reflect' ? '相邻单位图案先作轴对称变换，再平移' : config.motion === 'rotate' ? '相邻单位图案先旋转180°，再平移' : '各单位图案通过平移依次排列';
    return `以“${config.primaryName}${config.motifName}＋${config.secondaryName}${config.accentName}”组合作为单位图案，${direction}；${operation}，重复排列成二方连续纹样。创作参数：相邻单位间距 ${config.step}。`;
  }

  function validateLearningExplanation(text, row) {
    const value = String(text || '').trim();
    const rules = {
      '2': { object: /(蓝|曲线|纹样)/, motion: /平移/, forbidden: /(旋转|轴对称|反射)/, label: '第2行的蓝色纹样和平移' },
      '3': { object: /(黑白|几何|圆点|纹样)/, motion: /(轴对称|反射)/, also: /平移/, label: '第3行的黑白几何图案、轴对称和平移' },
      '4': { object: /(金|圆环|纹样)/, motion: /旋转/, also: /平移/, label: '第4行的金色纹样、旋转和平移' }
    };
    const rule = rules[String(row || '')];
    const missing = [];
    if (!rule) return { ok: false, missing: ['先选择要解释的教材纹样'] };
    if (value.length < 24) missing.push('至少24个字的完整说明');
    if (!/(单位图案|作为一个单位|一组完整)/.test(value) || !rule.object.test(value)) missing.push('说清这一行具体的单位图案');
    if (!rule.motion.test(value) || (rule.also && !rule.also.test(value)) || (rule.forbidden && rule.forbidden.test(value))) missing.push('与' + rule.label + '一致的变换方式');
    if (!/(左右|水平|向左|向右|横向)/.test(value)) missing.push('左右或水平排列方向');
    if (!/(连续|重复|依次|排列|生成)/.test(value)) missing.push('怎样继续排列');
    return { ok: missing.length === 0, missing };
  }

  function meaningful(value, minLength) {
    const text = String(value || '').trim();
    return text.length >= minLength && !/^(?:\d+|[a-z]+|[好行是无不]+)$/i.test(text);
  }

  function validateEvidence(config, evidence) {
    const problems = [];
    const unit = String(evidence.unitEvidence || '');
    if (!meaningful(unit, 10) || !(unit.includes(config.motifName) && unit.includes(config.accentName))) problems.push({ id: 'unitEvidence', label: `单位图案要写出“${config.motifName}”和“${config.accentName}”` });
    const motion = String(evidence.motionEvidence || '');
    const motionOk = config.motion === 'reflect' ? /(轴对称|反射).*平移|平移.*(轴对称|反射)/.test(motion) : config.motion === 'rotate' ? /旋转.*平移|平移.*旋转/.test(motion) : /平移/.test(motion);
    if (!meaningful(motion, 8) || !motionOk) problems.push({ id: 'motionEvidence', label: '变换说明要与你确认的规则一致' });
    const direction = String(evidence.directionEvidence || '');
    const directionOk = config.direction === 'left' ? /(左|水平)/.test(direction) : config.direction === 'both' ? /(左右|两侧|双向|水平)/.test(direction) : /(右|水平)/.test(direction);
    if (!meaningful(direction, 8) || !directionOk) problems.push({ id: 'directionEvidence', label: '排列方向要与图中箭头一致' });
    if (String(evidence.peerCode || '').trim().length < 2) problems.push({ id: 'peerCode', label: '填写实际续画的同伴编号' });
    if (!['yes', 'partly', 'no'].includes(evidence.peerResult)) problems.push({ id: 'peerResult', label: '选择同伴续画结果' });
    if (!meaningful(evidence.peerNote, 12)) problems.push({ id: 'peerNote', label: '具体记录同伴画对或需修改的地方' });
    const transfer = String(evidence.transferAnswer || '');
    if (!meaningful(transfer, 24) || !/(单位|图案)/.test(transfer) || !/(左右|水平|向左|向右)/.test(transfer) || !/(平移|轴对称|反射|旋转)/.test(transfer)) problems.push({ id: 'transferAnswer', label: '迁移回答要完整写出单位图案、方向和变换方式' });
    return problems;
  }

  const api = { interpretIdea, motifShape, motionTransform, ruleSentence, validateLearningExplanation, validateEvidence, meaningful, esc, DEFAULT, MOTIFS, COLOR_WORDS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.YunbaoStudio = api;
})(typeof window !== 'undefined' ? window : globalThis);
