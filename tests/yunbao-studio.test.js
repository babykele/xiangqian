const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const studio = require('../网页实验/yunbao-studio.js');

{
  const r = studio.interpretIdea('用蓝色云朵和黄色星星，一正一反向右重复，排列紧密，看起来活泼');
  assert.equal(r.motif, 'cloud');
  assert.equal(r.accent, 'star');
  assert.equal(r.primary, '#2878D0');
  assert.equal(r.secondary, '#F4B942');
  assert.equal(r.motion, 'reflect');
  assert.equal(r.step, 64);
  assert.equal(r.mood, '活泼');
  assert.ok(r.recognized.some(x => x.includes('生成方向')));
}

{
  const r = studio.interpretIdea('我想用绿色星星和红色叶子组成单位图案，向左连续排列，相邻图案旋转180度');
  assert.equal(r.motif, 'star');
  assert.equal(r.accent, 'leaf');
  assert.equal(r.primary, '#39A878');
  assert.equal(r.secondary, '#E45757');
  assert.equal(r.direction, 'left');
  assert.equal(r.motion, 'rotate');
}

{
  const r = studio.interpretIdea('紫色菱形旋转180度，每次平移90，有未来感');
  assert.equal(r.motif, 'diamond');
  assert.equal(r.primary, '#7B61C9');
  assert.equal(r.motion, 'rotate');
  assert.equal(r.step, 90);
  assert.equal(r.mood, '未来感');
}

{
  const r = studio.interpretIdea('我想做一件很好看的衣服');
  assert.ok(r.assumptions.length >= 3);
  assert.equal(r.motif, 'cloud');
  assert.equal(r.motion, 'translate');
}

{
  const sentence = studio.ruleSentence({...studio.DEFAULT, motion: 'reflect', direction: 'both', step: 64});
  assert.match(sentence, /沿左右方向连续排列/);
  assert.match(sentence, /轴对称变换/);
  assert.match(sentence, /64/);
  assert.match(sentence, /单位图案/);
  assert.match(sentence, /云蓝云朵/);
  assert.match(sentence, /暖黄星星/);
  assert.doesNotMatch(sentence, /形状和大小保持不变/);
}

{
  const weak = studio.validateLearningExplanation('单位图案沿左右方向平移，连续重复排列成二方连续纹样。', '3');
  assert.equal(weak.ok, false);
  const strong = studio.validateLearningExplanation('把一个完整的黑白几何图案作为单位图案，相邻单位先作轴对称再平移，沿左右方向连续排列。', '3');
  assert.equal(strong.ok, true);
}

{
  const config = studio.interpretIdea('用绿色星星和红色叶子组成单位图案，向左连续排列，相邻图案旋转180度');
  const bad = studio.validateEvidence(config, {unitEvidence:'1', motionEvidence:'1', directionEvidence:'1', peerCode:'1', peerResult:'', peerNote:'1', transferAnswer:'1'});
  assert.ok(bad.length >= 6);
  const good = studio.validateEvidence(config, {
    unitEvidence:'一个完整单位图案由星星和叶片组成。',
    motionEvidence:'相邻单位先旋转180度，再平移到下一个位置。',
    directionEvidence:'从起始单位沿水平方向向左连续排列。',
    peerCode:'07', peerResult:'yes', peerNote:'同伴画对了旋转方向和平移位置，可以独立续画。',
    transferAnswer:'取星星和叶片组成的单位图案，沿水平方向向左排列，每次先旋转180度再平移，就能画得更长。'
  });
  assert.deepEqual(good, []);
}

assert.match(studio.motifShape('cloud', '#fff', '#000'), /path/);
assert.match(studio.motifShape('star', '#fff', '#000'), /M0 -18/);
assert.equal(studio.motionTransform('reflect', 1), 'scale(-1 1)');
assert.equal(studio.motionTransform('rotate', 2), '');
assert.equal(studio.esc('<script>'), '&lt;script&gt;');

{
  const page = fs.readFileSync(path.join(__dirname, '../网页实验/云宝二方连续纹样设计工坊_v1.html'), 'utf8');
  const learn = page.indexOf('1 学习纹样');
  const speak = page.indexOf('2 说出创意');
  const make = page.indexOf('3 平台实现');
  assert.ok(learn >= 0 && learn < speak && speak < make);
  assert.match(page, /function checkExamples\(\)/);
  assert.match(page, /function finishLearning\(\)/);
  assert.match(page, /assets\/云宝吉祥物\.png/);
  assert.match(page, /yunbao-frieze-studio-v1-4/);
  assert.match(page, /step===4&&!S\.designConfirmed/);
  assert.match(page, /teacherReview:'pending'/);
  assert.match(page, /仅保存在本机，尚未提交给教师/);
  assert.match(page, /教材学习区/);
  assert.match(page, /迁移创作区/);
  assert.match(page, /均直接取自教材第137页表1/);
  assert.match(page, /教材表1_平移纹样\.png/);
  assert.match(page, /教材表1_轴对称平移纹样\.png/);
  assert.match(page, /教材表1_旋转平移纹样\.png/);
  assert.match(page, /颜色与间距是创作参数，不作为教材知识考查/);
  assert.match(page, /云宝二方连续纹样 AI 实验平台/);
  assert.match(page, /AI协同 · 智能体追问＋本地规则执行/);
  assert.match(page, /function openAiCoach\(\)/);
  assert.match(page, /ddd62b0d-7443-478b-91e4-444d0fb60c69/);
  assert.match(page, /AI对话与画布暂不自动同步，必须由你核对/);
  assert.doesNotMatch(page, /云朵和星星组合平移排列/);
  assert.doesNotMatch(page, /叶片交替反射排列/);
  assert.doesNotMatch(page, /最小单位纹样/);
  assert.doesNotMatch(page, /用数学证明你的设计/);
  for (const name of ['教材表1_示范纹样.png','教材表1_平移纹样.png','教材表1_轴对称平移纹样.png','教材表1_旋转平移纹样.png']) {
    assert.ok(fs.existsSync(path.join(__dirname, '../网页实验/assets/textbook', name)), `missing textbook asset: ${name}`);
  }
}

console.log('PASS yunbao studio: textbook boundary, gated learning, mascot asset, interpretation, rules and SVG motifs.');
