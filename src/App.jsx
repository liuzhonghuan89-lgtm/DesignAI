import React, { useState, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx'; 
import { motion, AnimatePresence } from 'framer-motion'; 
import { 
  Sparkles, ShoppingBag, Box, Layout, Palette, 
  Type, Image as ImageIcon, Layers, ChevronLeft, 
  Check, Wand2, X, ZoomIn, Droplet, Globe, BookOpen, 
  AlignLeft, Grid3X3, Armchair, Settings, Save, Trash2, Plus,
  FileSpreadsheet, BrainCircuit
} from 'lucide-react';

export default function DesignPlatform() {
  // --- 状态管理 ---
  const [step, setStep] = useState(1);
  const [loading, setLoading] = useState(false);
  const [loadingText, setLoadingText] = useState(""); 
  const [dbData, setDbData] = useState([]); 
  const [isDbLoaded, setIsDbLoaded] = useState(false);
  
  const [selectedImage, setSelectedImage] = useState(null);
  const [copyTip, setCopyTip] = useState(null); 
  const [showEditor, setShowEditor] = useState(false);
  const [showSmartPaste, setShowSmartPaste] = useState(false); 
  const [pasteText, setPasteText] = useState(""); 
  const [isParsing, setIsParsing] = useState(false); 

  const fileInputRef = useRef(null); 

  const [selections, setSelections] = useState({
    platform: '',
    site: '', 
    category: '',
    type: '',
    style: ''
  });

  const [finalSpec, setFinalSpec] = useState(null);

  const TEXT_KEY = import.meta.env.VITE_DASHSCOPE_KEY;
  const IMAGE_KEY = import.meta.env.VITE_NANO_KEY; 

  // 初始化：加载 CSV
  const loadData = () => {
    Papa.parse(`/specs.csv?t=${new Date().getTime()}`, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const cleanData = results.data.filter(row => row.platform && row.category);
        setDbData(cleanData);
        setIsDbLoaded(true);
      },
      error: () => alert("无法读取 public/specs.csv")
    });
  };

  useEffect(() => {
    loadData();
  }, []);

  // --- 核心排版渲染函数 ---
  const renderFormattedContent = (text) => {
    if (!text) return null;
    const rawText = text.replace(/\\n/g, '\n');
    const paragraphs = rawText.split('\n').filter(p => p.trim() !== '');
    
    return paragraphs.map((paragraph, idx) => {
      const parts = paragraph.split(/(\*\*.*?\*\*)/g);
      return (
        <div key={idx} className="mb-4 last:mb-0 leading-relaxed text-justify"> 
          {parts.map((part, i) => {
            if (part.startsWith('**') && part.endsWith('**')) {
              return <strong key={i} className="text-slate-900 font-extrabold">{part.slice(2, -2)}</strong>;
            }
            return <span key={i} className="text-slate-600">{part}</span>;
          })}
        </div>
      );
    });
  };

  // --- 数据保存与处理 ---
  const handleSaveData = async (newData) => {
    const csv = Papa.unparse(newData);
    try {
      const response = await fetch('/api/save-csv', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ csvContent: csv })
      });

      if (response.ok) {
        setDbData(newData);
        alert("✅ 数据已保存并同步到服务器！");
        setShowEditor(false);
        loadData();
        setStep(1);
      } else {
        alert("❌ 保存失败，请检查终端报错");
      }
    } catch (error) {
      console.error(error);
      alert("保存出错: " + error.message);
    }
  };

  const handleExcelImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);
        if (data.length > 0 && !data[0].platform) {
          alert("导入失败：Excel 表头必须包含 'platform', 'site', 'category' 等英文列名");
          return;
        }
        if (window.confirm(`成功读取 ${data.length} 条数据，确定追加吗？`)) {
          setDbData([...dbData, ...data]);
          alert("已追加数据，请点击‘保存并同步’！");
        }
      } catch (err) {
        console.error(err);
        alert("文件读取失败");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = null;
  };

  const handleSmartParse = async () => {
    if (!pasteText.trim()) return;
    setIsParsing(true);
    const smartPrompt = `
      你是一个数据清洗专家。请分析用户输入，提取CSV数据行。
      目标字段: platform, site, category, type, style, target_audience, colors, props
      用户输入: ${pasteText}
      要求: 1. 尽可能推断缺失字段。 2. 返回纯JSON数组。
    `;
    try {
      const response = await fetch("https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions", {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${TEXT_KEY}` },
        body: JSON.stringify({ model: "qwen-plus", messages: [{ role: "user", content: smartPrompt }] })
      });
      const resData = await response.json();
      const content = resData.choices[0].message.content.replace(/```json|```/g, '').trim();
      const parsedData = JSON.parse(content);
      if (Array.isArray(parsedData)) {
        setDbData([...dbData, ...parsedData]);
        setShowSmartPaste(false);
        setPasteText("");
        alert(`成功识别 ${parsedData.length} 条数据！请保存。`);
      } else { alert("识别失败"); }
    } catch (err) { console.error(err); alert("AI解析失败"); } finally { setIsParsing(false); }
  };

  const getOptionsForStep = (currentStep) => {
    let filtered = dbData;
    if (currentStep > 1) filtered = filtered.filter(row => row.platform === selections.platform);
    if (currentStep > 2) filtered = filtered.filter(row => row.site === selections.site);
    if (currentStep > 3) filtered = filtered.filter(row => row.category === selections.category);
    if (currentStep > 4) filtered = filtered.filter(row => row.type === selections.type);

    let key = '';
    if (currentStep === 1) key = 'platform';
    else if (currentStep === 2) key = 'site';
    else if (currentStep === 3) key = 'category';
    else if (currentStep === 4) key = 'type';
    else if (currentStep === 5) key = 'style';

    const options = [...new Set(filtered.map(row => row[key]))];
    return options.filter(opt => opt && opt.trim() !== '');
  };

  const handleSelect = (key, value) => {
    setSelections(prev => ({ ...prev, [key]: value }));
    if (step < 6) {
      setTimeout(() => setStep(step + 1), 250);
    }
  };

  const handleBack = () => {
    if (step > 1) setStep(step - 1);
  };

  const handleCopyColor = (hex, e) => {
    e.stopPropagation();
    navigator.clipboard.writeText(hex);
    setCopyTip(`色值 ${hex} 已复制`);
    setTimeout(() => setCopyTip(null), 2000);
  };

  const getLanguageInstruction = (site) => {
    const s = site.toLowerCase();
    if (s.includes('jp') || s.includes('japan')) return "Japanese (Katakana/Kanji) typography";
    if (s.includes('kr') || s.includes('korea')) return "Korean (Hangul) typography";
    if (s.includes('cn') || s.includes('china')) return "Chinese typography";
    if (s.includes('de') || s.includes('germany')) return "German typography";
    if (s.includes('fr') || s.includes('france')) return "French typography";
    return "English typography"; 
  };

  // --- API 核心逻辑 ---
  const generateDesign = async () => {
    if (!TEXT_KEY || !IMAGE_KEY) {
      alert("请检查 .env 文件，确保密钥已填写！");
      return;
    }

    const matchedRow = dbData.find(row => 
      row.platform === selections.platform &&
      row.site === selections.site &&
      row.category === selections.category &&
      row.type === selections.type &&
      row.style === selections.style
    );

    if (!matchedRow) {
      alert("未找到匹配的规范数据，请检查 CSV 是否完整。");
      setStep(6);
      return;
    }

    setLoading(true);
    setStep(7);

    try {
      setLoadingText("AI 正在根据站点文化解构策略...");
      const aiData = await callTextAI(matchedRow, selections.site);
      
      setLoadingText("AI 正在渲染本地化视觉方案...");
      
      // 容错处理：如果 AI 没生成 scenePrompt，用兜底词
      const baseScenePrompt = aiData.scenePrompt || `High quality product photography of ${matchedRow.category}, ${matchedRow.style} style, 8k resolution`;
      const colorPrompt = aiData.palette ? `Color palette: ${aiData.palette.map(c => c.name).join(', ')}` : "";
      const langPrompt = getLanguageInstruction(selections.site); 

      // 🔴 关键修复：并发生成时，确保 Prompt 不会因为太长而报错
      const [sceneImage, adImage] = await Promise.all([
        callImageAI(matchedRow, `Product photography, minimal clean background, no text, ${colorPrompt}, ${baseScenePrompt}`),
        callImageAI(matchedRow, `Commercial poster design, ${langPrompt}, elegant typography layout, headline text placeholder, ${colorPrompt}, ${baseScenePrompt}`)
      ]);

      setFinalSpec({
        ...matchedRow,
        aiData: aiData, 
        images: [sceneImage, adImage] 
      });

    } catch (error) {
      console.error(error);
      alert("生成失败: " + error.message);
      setStep(6);
    } finally {
      setLoading(false);
    }
  };

  const callTextAI = async (row, site) => {
    const prompt = `
      你是一个拥有10年经验的跨境电商视觉总监。请根据以下规范，为设计师生成一份**极具执行落地性、深度且专业**的《视觉设计执行手册 (SOP)》。
      
      [项目数据] 
      平台: ${row.platform}
      站点: ${site}
      风格: ${row.style}
      受众: ${row.target_audience}
      配色: ${row.colors}
      道具: ${row.props}
      
      [核心要求]
      1. **深度分析**：结合 ${site} 站点的文化背景解释“为什么要这样做”。
      2. **参数量化**：涉及字体、间距、占比时，必须给出具体数值（px, pt, %）。
      3. **排版格式**：**关键：每个独立的建议点（如主标题、正文、层级）之间必须使用换行符 \\n 分隔。** 必须使用 **加粗** 标记关键参数。
      4. **内容详实**：每个板块的内容不少于 80 字，拒绝空洞。
      
      请严格按照以下 JSON 格式返回 (不要 Markdown 代码块):
      {
        "fonts": { "primary": "主字体名", "secondary": "辅助字体名" },
        "palette": [
           {"hex": "#xxxxxx", "name": "颜色名", "ratio": 40}
        ],
        "layout_keywords": ["关键词1", "关键词2"],
        "scenePrompt": "English DALL-E prompt...",
        "design_guide": [
           {
             "title": "字体规范", 
             "content": "**主标题选型：**强烈推荐使用 **Helvetica Neue Bold**，字号需设定在 **48-60pt** 之间以确保冲击力。\\n**正文排版：**建议使用 **Roboto Regular (16pt)**，行高设定为 **1.6倍**，符合${site}地区阅读习惯。\\n**层级对比：**标题与正文的字重对比度应达到 **3:1**，确保信息层级清晰。"
           },
           {
             "title": "配色策略", 
             "content": "**主色逻辑：**主色 **#Hex值** 应占据画面 **40%** 面积，重点用于产品主体及 CTA 按钮。\\n**辅助色应用：**辅助色主要用于背景氛围渲染，建议使用低饱和度色系，占比 **30%**。\\n**文化禁忌：**严禁使用过于刺眼的荧光色，以免产生廉价感。"
           },
           {
             "title": "构图形式", 
             "content": "**构图法则：**严格遵循 **三分法**，将产品主体置于视觉中心的 **左侧 2/3 区域**。\\n**留白策略：**画面 **右上角需保留 30%** 的纯净留白区域，用于放置文案。\\n**视线引导：**利用光影形成 **Z型** 视线引导，聚焦核心卖点。"
           },
           {
             "title": "道具摆放", 
             "content": "**空间关系：**道具必须严格置于产品 **后方 20cm** 处，切勿与产品处于同一焦平面。\\n**材质呼应：**道具材质建议选择 **哑光/磨砂** 质感，与产品形成对比。\\n**场景叙事：**道具选择必须服务于使用场景，暗示高效舒适的氛围。"
           }
        ]
      }
    `;
    const url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions";
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${TEXT_KEY}` },
        body: JSON.stringify({ model: "qwen-plus", messages: [{ role: "user", content: prompt }] })
      });
      if (!response.ok) throw new Error("文本服务繁忙");
      const data = await response.json();
      const text = data.choices[0].message.content.replace(/```json|```/g, '').trim();
      return JSON.parse(text);
    } catch (err) { throw err; }
  };

  // 🔴 修复核心：增加 prompt 截断逻辑，防止 Render Failed
  const callImageAI = async (row, promptSuffix) => {
    const url = "https://www.dmxapi.cn/v1/images/generations";
    const typeName = selections.type.toLowerCase();
    let aspect_ratio = "1:1";
    if (typeName.includes("高级") || typeName.includes("21:9")) aspect_ratio = "21:9";
    else if (typeName.includes("a+") || typeName.includes("16:9")) aspect_ratio = "16:9";
    else if (typeName.includes("竖版") || typeName.includes("3:4")) aspect_ratio = "3:4";
    else if (typeName.includes("主图") || typeName.includes("1:1")) aspect_ratio = "1:1";

    // 1. 构建完整 Prompt
    let fullPrompt = `${row.category}, ${row.style} style. ${promptSuffix}`;
    
    // 2. 🚨 安全截断：防止 Prompt 过长导致 API 报错 (通常限制在 1000 字符以内)
    if (fullPrompt.length > 900) {
      console.warn("Prompt too long, truncating...", fullPrompt.length);
      fullPrompt = fullPrompt.substring(0, 900); // 截取前900个字符
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${IMAGE_KEY}` },
        body: JSON.stringify({
          model: "nano-banana-2", 
          prompt: fullPrompt, // 使用截断后的 Prompt
          n: 1, size: "1k", 
          aspect_ratio: aspect_ratio, 
          response_format: "url"
        })
      });
      
      if (!response.ok) {
         // 打印详细错误方便调试
         const errData = await response.json();
         console.error("Image Gen Error:", errData);
         throw new Error("绘图服务繁忙");
      }
      
      const data = await response.json();
      const imgData = data.data[0];
      if (imgData.b64_json) return `data:image/png;base64,${imgData.b64_json}`;
      if (imgData.url) return imgData.url;
      throw new Error("无图片数据");
    } catch (err) {
      console.error(err);
      return `https://placehold.co/1024x1024/F3F4F6/9CA3AF?text=Render+Failed`;
    }
  };

  const bgPattern = {
    backgroundImage: `radial-gradient(#94A3B8 1px, transparent 1px)`,
    backgroundSize: '32px 32px',
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-slate-800 font-sans selection:bg-slate-900 selection:text-white overflow-hidden relative flex flex-col" style={bgPattern}>
      
      <div className="fixed top-0 left-0 w-full h-40 bg-gradient-to-b from-[#F8F9FA] via-[#F8F9FA]/80 to-transparent pointer-events-none z-0"></div>

      <AnimatePresence>
        {copyTip && (
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className="fixed top-8 left-1/2 -translate-x-1/2 z-[60] bg-slate-900 text-white px-6 py-2 rounded-full shadow-lg text-sm font-bold flex items-center gap-2"
          >
            <Check size={16} className="text-green-400"/> {copyTip}
          </motion.div>
        )}
      </AnimatePresence>

      <div className="relative z-10 max-w-5xl mx-auto px-6 py-8 flex flex-col h-screen">
        
        {/* Header Grid */}
        <header className="flex-none w-full mb-8 h-20 grid grid-cols-3 items-center z-20 relative">
          <div className="flex items-center gap-3 justify-self-start">
             <div className="w-10 h-10 bg-slate-900 rounded-xl flex items-center justify-center shadow-lg shadow-slate-900/20">
                <Wand2 className="text-white w-5 h-5" />
             </div>
             <h1 className="text-2xl font-bold tracking-tight text-slate-900">DesignFlow</h1>
          </div>
          
          <div className="flex gap-2 bg-white/50 backdrop-blur-sm p-1.5 rounded-full border border-white/60 shadow-sm justify-self-center">
            {[1, 2, 3, 4, 5, 6, 7].map(i => (
              <motion.div 
                key={i} 
                initial={false}
                animate={{ 
                  width: step >= i ? 32 : 8,
                  backgroundColor: step >= i ? '#0F172A' : '#E2E8F0'
                }}
                className="h-2 rounded-full"
              />
            ))}
          </div>

          <div className="justify-self-end">
            <button 
              onClick={() => setShowEditor(true)}
              className="flex items-center gap-2 px-4 py-2 rounded-full bg-white border border-slate-200 text-slate-600 hover:text-slate-900 hover:border-slate-400 shadow-sm text-sm font-bold transition-all"
            >
              <Settings size={16} /> 数据管理
            </button>
          </div>
        </header>

        {/* Main Content */}
        <div className="flex-1 flex flex-col items-center justify-center relative w-full perspective-1000 min-h-0"> 
          <AnimatePresence mode='wait'>
            
            {/* Step 1-5: Selectors */}
            {step <= 5 && isDbLoaded && (
               <GlassCard key="step-selector" className="w-full max-w-4xl p-12 my-auto">
                  <motion.h2 
                    key={`title-${step}`}
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="text-3xl font-bold mb-12 text-center text-slate-800 flex items-center justify-center gap-3"
                  >
                     {step === 1 && <><ShoppingBag className="text-slate-400"/> 选择发布平台</>}
                     {step === 2 && <><Globe className="text-slate-400"/> 选择站点</>}
                     {step === 3 && <><Box className="text-slate-400"/> 选择产品类目</>}
                     {step === 4 && <><Layout className="text-slate-400"/> 选择设计类型</>}
                     {step === 5 && <><Palette className="text-slate-400"/> 选择设计风格</>}
                  </motion.h2>
                  
                  <div className="flex flex-wrap justify-center gap-5">
                    {getOptionsForStep(step).map((opt, index) => (
                      <motion.button
                        key={opt}
                        initial={{ opacity: 0, y: 20 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: index * 0.05 }}
                        whileHover={{ scale: 1.05, y: -2 }}
                        whileTap={{ scale: 0.95 }}
                        onClick={() => handleSelect(
                          step === 1 ? 'platform' : 
                          step === 2 ? 'site' : 
                          step === 3 ? 'category' : 
                          step === 4 ? 'type' : 'style', 
                          opt
                        )}
                        className="px-10 py-6 rounded-2xl bg-white border border-slate-100 shadow-[0_2px_10px_rgba(0,0,0,0.03)] hover:shadow-[0_8px_30px_rgba(0,0,0,0.08)] hover:border-slate-300 transition-all min-w-[180px] text-lg font-bold text-slate-700 relative overflow-hidden group"
                      >
                        <span className="relative z-10">{opt}</span>
                        <div className="absolute inset-0 bg-gradient-to-r from-transparent via-white/40 to-transparent -translate-x-full group-hover:animate-shimmer" />
                      </motion.button>
                    ))}
                  </div>

                  {step > 1 && (
                    <motion.div 
                      initial={{ opacity: 0 }} 
                      animate={{ opacity: 1 }}
                      className="mt-12 flex justify-center border-t border-slate-100 pt-8"
                    >
                      <button onClick={handleBack} className="flex items-center gap-2 px-8 py-3 rounded-full bg-slate-900 text-white text-sm font-bold shadow-lg hover:bg-black hover:scale-105 active:scale-95 transition-all">
                        <ChevronLeft size={16} /> 上一步
                      </button>
                    </motion.div>
                  )}
               </GlassCard>
            )}

            {/* Step 6: Confirmation */}
            {step === 6 && (
               <GlassCard key="confirm" className="w-full max-w-md p-10 text-center my-auto">
                 <motion.div initial={{ scale: 0 }} animate={{ scale: 1 }} className="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center mx-auto mb-6 shadow-inner"><Sparkles className="w-10 h-10 text-slate-900" /></motion.div>
                 <h3 className="text-2xl font-bold mb-2">准备生成</h3>
                 <p className="text-slate-500 text-sm mb-8">AI 将根据 {selections.site} 站点进行本地化设计</p>
                 <div className="space-y-3 mb-10 text-left">
                   <SummaryRow label="平台" value={selections.platform} delay={0.1} />
                   <SummaryRow label="站点" value={selections.site} delay={0.15} /> 
                   <SummaryRow label="类目" value={selections.category} delay={0.2} />
                   <SummaryRow label="类型" value={selections.type} delay={0.3} />
                   <SummaryRow label="风格" value={selections.style} delay={0.4} />
                 </div>
                 <div className="space-y-4">
                    <motion.button whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }} onClick={generateDesign} className="w-full py-5 text-lg font-bold text-white rounded-2xl bg-slate-900 hover:bg-black shadow-xl shadow-slate-900/20 transition-all">开始生成方案</motion.button>
                    <button onClick={handleBack} className="w-full py-3 text-sm font-bold text-slate-500 hover:text-slate-800 transition-colors">返回修改</button>
                 </div>
               </GlassCard>
            )}

            {/* Step 7: Result Page */}
            {step === 7 && (
              <motion.div 
                key="result"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 1.05 }}
                className="w-full h-full overflow-y-auto pr-2 pb-20" 
              >
                {loading ? (
                  <GlassCard className="flex flex-col items-center justify-center h-full min-h-[400px]">
                    <div className="relative w-24 h-24">
                      <motion.div className="absolute inset-0 border-4 border-slate-200 rounded-full" />
                      <motion.div className="absolute inset-0 border-4 border-slate-900 rounded-full border-t-transparent" animate={{ rotate: 360 }} transition={{ duration: 1, repeat: Infinity, ease: "linear" }} />
                    </div>
                    <h2 className="text-xl font-bold mt-8 text-slate-800">AI 正在设计中...</h2>
                    <motion.p key={loadingText} initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="text-slate-400 mt-2 text-sm font-medium">{loadingText}</motion.p>
                  </GlassCard>
                ) : finalSpec ? (
                  <div className="space-y-6 pb-20"> 
                    <GlassCard className="flex flex-col md:flex-row justify-between items-center p-8 sticky top-0 z-10 backdrop-blur-xl bg-white/90 border-b border-slate-200 shadow-sm">
                      <div>
                        <h2 className="text-4xl font-bold text-slate-900 tracking-tight">{finalSpec.style}</h2>
                        <div className="flex items-center gap-3 mt-3">
                           <span className="bg-slate-100 border border-slate-200 px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider text-slate-600">Site: {selections.site}</span>
                           <span className="text-slate-500 text-sm font-medium">{finalSpec.target_audience}</span>
                        </div>
                      </div>
                      <motion.button whileHover={{ scale: 1.05 }} whileTap={{ scale: 0.95 }} onClick={() => setStep(1)} className="mt-6 md:mt-0 px-8 py-3 rounded-full text-sm font-bold bg-white border border-slate-200 shadow-sm hover:shadow-md text-slate-900">开启新项目</motion.button>
                    </GlassCard>

                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                      <div className="space-y-6">
                        <InfoCard title="Typography / 字体策略" icon={<Type size={18}/>} delay={0.1}>
                           <div className="space-y-3">
                             <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                               <span className="text-[10px] text-slate-400 font-bold uppercase block mb-1">Primary Font</span>
                               <span className="text-lg font-bold text-slate-800">{finalSpec.aiData?.fonts?.primary || "Sans-Serif"}</span>
                             </div>
                             <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                               <span className="text-[10px] text-slate-400 font-bold uppercase block mb-1">Secondary Font</span>
                               <span className="text-md font-medium text-slate-700">{finalSpec.aiData?.fonts?.secondary || "Serif"}</span>
                             </div>
                           </div>
                        </InfoCard>
                        <InfoCard title="Layout / 构图关键词" icon={<Layout size={18}/>} delay={0.2}>
                           <div className="flex flex-wrap gap-2">
                             {finalSpec.aiData?.layout_keywords?.map((kw, i) => (
                               <span key={i} className="px-3 py-1.5 bg-slate-100 text-slate-600 text-xs font-bold rounded-lg border border-slate-200">#{kw}</span>
                             )) || <span className="text-slate-400 text-sm">暂无关键词</span>}
                           </div>
                        </InfoCard>
                        <InfoCard title="Props / 道具建议" icon={<Box size={18}/>} delay={0.3}>
                           <p className="text-slate-600 text-sm leading-relaxed">{finalSpec.props}</p>
                        </InfoCard>
                      </div>

                      <div className="lg:col-span-2 space-y-6">
                        <GlassCard className="p-8" delay={0.3}>
                           <h3 className="flex items-center gap-2 text-lg font-bold mb-8 text-slate-900 border-b border-slate-100 pb-4">
                             <ImageIcon size={20}/> 视觉方案 & 智能色卡
                           </h3>
                           <div className="grid md:grid-cols-2 gap-8 mb-10">
                              <ResultImage src={finalSpec.images[0]} tag="Scene" desc="纯净场景参考" delay={0.4} onZoom={() => setSelectedImage(finalSpec.images[0])}/>
                              <ResultImage src={finalSpec.images[1]} tag="Layout" desc="商业排版参考" delay={0.5} onZoom={() => setSelectedImage(finalSpec.images[1])}/>
                           </div>
                           <div className="bg-slate-50 rounded-2xl p-6 border border-slate-100">
                              <div className="flex items-center justify-between mb-4">
                                 <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2"><Droplet size={16} className="text-slate-400"/> AI 推荐配色 (点击复制)</h4>
                                 <span className="text-[10px] text-slate-400 font-bold uppercase">Local Culture Optimized</span>
                              </div>
                              <div className="flex h-16 rounded-xl overflow-hidden shadow-sm border border-slate-200 ring-4 ring-white">
                                 {finalSpec.aiData?.palette?.map((color, idx) => (
                                    <div key={idx} onClick={(e) => handleCopyColor(color.hex, e)} className="h-full relative group cursor-pointer hover:brightness-110" style={{ backgroundColor: color.hex, width: `${color.ratio}%` }}>
                                      <div className="absolute inset-0 flex flex-col items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity duration-200 bg-black/20 backdrop-blur-[2px]">
                                         <span className="text-white text-xs font-bold drop-shadow-md">{color.hex}</span>
                                         <span className="text-white/80 text-[10px] font-medium">{color.ratio}%</span>
                                      </div>
                                    </div>
                                 ))}
                              </div>
                           </div>
                        </GlassCard>

                        <GlassCard className="p-8 relative overflow-hidden" delay={0.4}>
                           <div className="absolute top-0 right-0 p-32 bg-slate-100 rounded-full blur-[80px] opacity-50 -z-10 pointer-events-none"></div>
                           <h3 className="flex items-center gap-2 text-lg font-bold mb-6 text-slate-900 border-b border-slate-100 pb-4">
                             <BookOpen size={20}/> AI 设计执行手册 (SOP)
                           </h3>
                           <div className="grid md:grid-cols-2 gap-6">
                              {Array.isArray(finalSpec.aiData?.design_guide) ? (
                                finalSpec.aiData.design_guide.map((item, index) => (
                                  <div key={index} className="bg-white/60 p-5 rounded-xl border border-slate-100 hover:shadow-md transition-shadow">
                                     <h4 className="flex items-center gap-2 font-bold text-slate-800 mb-3 text-sm">
                                        <span className="p-1.5 bg-slate-100 rounded-md text-slate-500">
                                          {item.title.includes("字体") ? <Type size={14}/> : 
                                           item.title.includes("配色") ? <Palette size={14}/> :
                                           item.title.includes("构图") ? <Grid3X3 size={14}/> :
                                           item.title.includes("道具") ? <Armchair size={14}/> : <AlignLeft size={14}/>}
                                        </span>
                                        {item.title}
                                     </h4>
                                     <div className="text-sm text-slate-600 leading-relaxed text-justify">
                                        {renderFormattedContent(item.content)}
                                     </div>
                                  </div>
                                ))
                              ) : (
                                <div className="md:col-span-2 prose prose-slate max-w-none text-sm text-slate-600 leading-7 whitespace-pre-line font-medium bg-white/50 p-6 rounded-2xl border border-slate-100/50">
                                   {finalSpec.aiData?.design_guide || "AI 正在撰写执行建议..."}
                                </div>
                              )}
                           </div>
                        </GlassCard>
                      </div>
                    </div>
                  </div>
                ) : null}
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Editor & Smart Paste & Lightbox - 保持不变 */}
        <AnimatePresence>
          {showEditor && (
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
              <motion.div initial={{ scale: 0.95, y: 20 }} animate={{ scale: 1, y: 0 }} exit={{ scale: 0.95, y: 20 }} className="bg-white w-full max-w-6xl h-[90vh] rounded-2xl shadow-2xl flex flex-col overflow-hidden">
                <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
                  <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><Settings size={20}/> 数据库管理 (specs.csv)</h2>
                  <div className="flex gap-3">
                    <button onClick={() => setShowSmartPaste(true)} className="flex items-center gap-2 bg-purple-100 text-purple-700 px-4 py-2 rounded-lg text-sm font-bold hover:bg-purple-200 transition-colors"><BrainCircuit size={16}/> AI 智能录入</button>
                    <div className="relative">
                      <input type="file" accept=".xlsx, .xls" className="hidden" ref={fileInputRef} onChange={handleExcelImport}/>
                      <button onClick={() => fileInputRef.current.click()} className="flex items-center gap-2 bg-green-100 text-green-700 px-4 py-2 rounded-lg text-sm font-bold hover:bg-green-200 transition-colors"><FileSpreadsheet size={16}/> 导入 Excel</button>
                    </div>
                    <button onClick={() => setShowEditor(false)} className="p-2 hover:bg-slate-200 rounded-full ml-4"><X size={20}/></button>
                  </div>
                </div>
                <div className="flex-1 overflow-auto p-6 bg-slate-50/50">
                  <table className="w-full text-sm text-left text-slate-600 border-collapse">
                    <thead className="text-xs text-slate-700 uppercase bg-slate-100 sticky top-0 z-10 shadow-sm">
                      <tr>{dbData.length > 0 && Object.keys(dbData[0]).map(key => (<th key={key} className="px-4 py-3 border-b border-slate-200 whitespace-nowrap">{key}</th>))}<th className="px-4 py-3 border-b border-slate-200 text-right sticky right-0 bg-slate-100">操作</th></tr>
                    </thead>
                    <tbody>
                      {dbData.map((row, index) => (
                        <tr key={index} className="bg-white border-b border-slate-100 hover:bg-blue-50/50 group">
                          {Object.keys(row).map(key => (
                            <td key={key} className="px-4 py-3 border-r border-slate-50 last:border-r-0 min-w-[120px] max-w-[240px]">
                              <input className="w-full bg-transparent border-none focus:ring-0 p-0 text-slate-700 text-sm truncate focus:text-clip" value={row[key]} title={row[key]} onChange={(e) => { const newData = [...dbData]; newData[index][key] = e.target.value; setDbData(newData); }} />
                            </td>
                          ))}
                          <td className="px-4 py-3 text-right sticky right-0 bg-white group-hover:bg-blue-50/50 shadow-[-10px_0_10px_-5px_rgba(0,0,0,0.05)]">
                            <button onClick={() => { const newData = dbData.filter((_, i) => i !== index); setDbData(newData); }} className="text-slate-400 hover:text-red-500 transition-colors p-1"><Trash2 size={16}/></button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <button onClick={() => { const newRow = Object.keys(dbData[0]).reduce((acc, key) => ({...acc, [key]: ''}), {}); setDbData([...dbData, newRow]); }} className="mt-4 flex items-center gap-2 text-blue-600 font-bold hover:bg-blue-50 px-4 py-2 rounded-lg transition-colors"><Plus size={18}/> 添加新行</button>
                </div>
                <div className="p-6 border-t border-slate-100 bg-white flex justify-end gap-3">
                  <button onClick={() => setShowEditor(false)} className="px-6 py-2.5 rounded-xl text-slate-600 font-bold hover:bg-slate-100 transition-colors">取消</button>
                  <button onClick={() => handleSaveData(dbData)} className="px-6 py-2.5 rounded-xl bg-slate-900 text-white font-bold hover:bg-black shadow-lg flex items-center gap-2"><Save size={18}/> 保存并同步</button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        <AnimatePresence>
          {showSmartPaste && (
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 z-[60] flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
              <motion.div initial={{ scale: 0.9, y: 20 }} animate={{ scale: 1, y: 0 }} exit={{ scale: 0.9, y: 20 }} className="bg-white w-full max-w-2xl rounded-2xl shadow-2xl p-8">
                <h3 className="text-2xl font-bold mb-4 flex items-center gap-2 text-slate-800"><BrainCircuit className="text-purple-600"/> AI 智能录入</h3>
                <p className="text-slate-500 mb-6 text-sm">直接粘贴包含产品信息的文本（如：亚马逊 Listing 描述、客户需求文档、聊天记录），AI 将自动提取并填入表格。</p>
                <textarea className="w-full h-48 p-4 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-purple-500 focus:border-transparent resize-none text-slate-700 mb-6" placeholder="在此粘贴文本..." value={pasteText} onChange={(e) => setPasteText(e.target.value)} />
                <div className="flex justify-end gap-3">
                  <button onClick={() => setShowSmartPaste(false)} className="px-6 py-3 rounded-xl text-slate-600 font-bold hover:bg-slate-100 transition-colors">取消</button>
                  <button onClick={handleSmartParse} disabled={isParsing || !pasteText.trim()} className="px-8 py-3 rounded-xl bg-purple-600 text-white font-bold hover:bg-purple-700 shadow-lg disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2">{isParsing ? '正在分析...' : '开始识别'}</button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        <AnimatePresence>
          {selectedImage && (
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setSelectedImage(null)} className="fixed inset-0 z-[70] flex items-center justify-center bg-black/95 backdrop-blur-xl cursor-zoom-out p-8">
              <motion.img initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} src={selectedImage} className="max-w-full max-h-full object-contain rounded-lg shadow-2xl" />
              <div className="absolute top-6 right-6 text-white/50 hover:text-white cursor-pointer"><X size={32} /></div>
            </motion.div>
          )}
        </AnimatePresence>

      </div>
    </div>
  );
}

// --- Component Definitions ---
const GlassCard = ({ children, className, delay = 0 }) => (
  <motion.div initial={{ opacity: 0, y: 20, scale: 0.98 }} animate={{ opacity: 1, y: 0, scale: 1 }} exit={{ opacity: 0, y: -20, scale: 0.98 }} transition={{ duration: 0.4, delay: delay, ease: "easeOut" }} className={`bg-white/80 backdrop-blur-2xl border border-white/60 shadow-[0_20px_40px_rgba(0,0,0,0.04)] rounded-[32px] ${className}`}>{children}</motion.div>
);
const SummaryRow = ({ label, value, delay }) => (
  <motion.div initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} transition={{ delay }} className="flex justify-between items-center py-3 px-4 bg-white rounded-xl border border-slate-100 shadow-sm"><span className="text-slate-400 text-xs font-bold uppercase tracking-wider">{label}</span><span className="font-bold text-slate-800 text-sm">{value}</span></motion.div>
);
const InfoCard = ({ title, icon, children, delay }) => (
  <GlassCard className="p-8" delay={delay}><h3 className="flex items-center gap-3 text-sm font-bold mb-6 text-slate-900 uppercase tracking-wider"><span className="p-2 bg-slate-100 rounded-lg text-slate-600">{icon}</span> {title}</h3>{children}</GlassCard>
);
const ResultImage = ({ src, tag, desc, delay, onZoom }) => (
  <motion.div initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay }} className="group space-y-4">
    <div onClick={onZoom} className="bg-slate-100 rounded-2xl overflow-hidden shadow-sm border border-slate-200 relative aspect-square hover:shadow-2xl hover:shadow-slate-200 transition-all duration-500 cursor-zoom-in"><img src={src} alt={tag} className="w-full h-full object-cover transition-transform duration-700 group-hover:scale-105" /><div className="absolute top-4 left-4"><div className="bg-white/90 backdrop-blur-md text-slate-900 text-[10px] font-extrabold px-3 py-1.5 rounded-full uppercase tracking-wide shadow-sm flex items-center gap-1"><div className={`w-2 h-2 rounded-full ${tag === 'Scene' ? 'bg-blue-500' : 'bg-purple-500'}`}></div>{tag}</div></div><div className="absolute inset-0 bg-black/0 group-hover:bg-black/10 transition-colors duration-300 flex items-center justify-center opacity-0 group-hover:opacity-100"><div className="bg-white/90 backdrop-blur text-slate-900 px-4 py-2 rounded-full font-bold shadow-lg transform translate-y-4 group-hover:translate-y-0 transition-all flex items-center gap-2"><ZoomIn size={16}/> 点击放大</div></div></div><div className="flex justify-between items-center px-1"><p className="text-xs text-slate-500 font-bold uppercase tracking-wide">{desc}</p><Check size={14} className="text-green-500 opacity-0 group-hover:opacity-100 transition-opacity"/></div>
  </motion.div>
);