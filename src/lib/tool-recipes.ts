export type ToolRecipe = {
  id: string;
  title: string;
  description: string;
  mode: 'agent' | 'ask';
  prompt: string;
};
export const toolRecipes: ToolRecipe[] = [
  {
    id: 'proofread',
    title: '校对与润色',
    description: '纠正表达，保留事实与引用',
    mode: 'agent',
    prompt:
      '请校对当前范围中的错别字、标点和不通顺的表达。只做必要的局部修改，保留数字、事实、引用和原意；不要整段重写仅需修改一个词的段落。',
  },
  {
    id: 'shorten',
    title: '精简表达',
    description: '去掉重复，不牺牲关键信息',
    mode: 'agent',
    prompt:
      '请精简当前范围的冗余表达，目标减少约 20% 的篇幅。优先保留结论、条件、数据和引用；无法安全压缩时保留原文，不要编造新内容。',
  },
  {
    id: 'translate',
    title: '中译英',
    description: '保留数字、术语和段落结构',
    mode: 'agent',
    prompt:
      '请把当前范围的中文翻译为自然、准确的英文，保留数字、引用、专有名词和段落结构。术语保持一致，歧义不要自行添加事实。',
  },
  {
    id: 'summary',
    title: '要点摘要',
    description: '仅在对话中总结，不修改原文',
    mode: 'ask',
    prompt:
      '请在对话中总结当前范围的核心观点、关键数据和待解决问题。明确区分原文结论与尚未提供的信息；不要修改文档。',
  },
  {
    id: 'logic',
    title: '逻辑审阅',
    description: '定位论据不足和结构问题',
    mode: 'ask',
    prompt:
      '请审阅当前范围的逻辑和结构，指出具体段落中的矛盾、跳步、论据不足或重复，并给出建议。只讨论原文可支持的内容，不修改文档，不声称已联网核验。',
  },
  {
    id: 'actions',
    title: '提取待办',
    description: '负责人和期限缺失时明确标注',
    mode: 'ask',
    prompt:
      '请从当前范围提取行动项，以任务、负责人、截止时间、来源段落四列展示。原文没有负责人或日期时写“未指定”，不要推测；不要修改文档。',
  },
  {
    id: 'table',
    title: '整理成表格',
    description: '先检查结构，避免虚构数据',
    mode: 'agent',
    prompt:
      '请将当前范围中适合对比的事实整理成一个简洁表格，插入相关段落之后并保留原文。缺失值写“未提供”，不要推测数据。选区模式只允许文字替换，请先在对话中说明需要切换到全文模式后再插表。',
  },
  {
    id: 'terms',
    title: '术语一致性',
    description: '列出疑似变体，先确认再改',
    mode: 'ask',
    prompt:
      '请检查当前范围中的术语、缩写和名称是否一致。列出疑似变体、出现位置及建议统一形式。相似词不一定同义，请标注不确定项；先给建议，不修改文档。',
  },
];
