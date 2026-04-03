import React, { useState, useRef, useEffect } from 'react';
import { FileDown, Type, Image as ImageIcon, Heading, Save, FolderOpen, MoveHorizontal, HelpCircle, X, MousePointerClick, Scissors, RefreshCw, ArrowDownToLine, Sparkles, History, BookOpen } from 'lucide-react';
import { Block, BlockType } from './types';
import { BlockRenderer } from './components/BlockRenderer';
import { exportToDocx } from './services/docxService';
import FileSaver from 'file-saver';

// Helper to generate IDs
const generateId = () => Math.random().toString(36).substr(2, 9);

// Helper: Convert Blob URL to Base64
const blobUrlToBase64 = async (blobUrl: string): Promise<string> => {
  // If it's already base64, return it
  if (blobUrl.startsWith('data:')) return blobUrl;
  
  try {
    const response = await fetch(blobUrl);
    if (!response.ok) throw new Error('Network response was not ok');
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.warn("Failed to fetch blob URL, it might be expired or from another session:", blobUrl);
    // Return a transparent 1x1 pixel base64 image as fallback
    return "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
  }
};

function App() {
  const [blocks, setBlocks] = useState<Block[]>([]);
  const [margin, setMargin] = useState(5);
  const [showHelp, setShowHelp] = useState(false);
  const [helpTab, setHelpTab] = useState<'guide' | 'updates'>('guide'); // 'guide' or 'updates'
  
  // Scale state for responsive A4 preview
  const [scale, setScale] = useState(1);
  const pageContainerRef = useRef<HTMLDivElement>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const jsonInputRef = useRef<HTMLInputElement>(null);
  
  // Drag and Drop State
  const [draggedBlockIndex, setDraggedBlockIndex] = useState<number | null>(null);
  const [targetBlockIndex, setTargetBlockIndex] = useState<number | null>(null); // Track where we are hovering
  const [editingBlockId, setEditingBlockId] = useState<string | null>(null);

  // --- Resize / Scale Logic ---
  useEffect(() => {
    const handleResize = () => {
        const A4_WIDTH_PX = 794; // Approx 210mm at 96 DPI
        const SCREEN_PADDING = 32; // 16px padding on each side
        const availableWidth = window.innerWidth - SCREEN_PADDING;

        if (availableWidth < A4_WIDTH_PX) {
            const newScale = availableWidth / A4_WIDTH_PX;
            setScale(newScale);
        } else {
            setScale(1);
        }
    };

    window.addEventListener('resize', handleResize);
    handleResize(); // Initial check

    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // --- Auto Count Logic ---
  const recalculateCounts = (currentBlocks: Block[]): Block[] => {
    let currentTextRowId: string | null = null;
    const counts: Record<string, number> = {};

    // Pass 1: Count images per Text Row
    for (const block of currentBlocks) {
      if (block.type === BlockType.TEXT_ROW) {
        currentTextRowId = block.id;
        counts[currentTextRowId] = 0;
      } else if (block.type === BlockType.IMAGE && currentTextRowId) {
        counts[currentTextRowId]++;
      }
    }

    // Pass 2: Update subContent
    return currentBlocks.map(block => {
      if (block.type === BlockType.TEXT_ROW) {
        const count = counts[block.id] || 0;
        const newSubContent = `好评${count}次`;
        if (block.subContent !== newSubContent) {
          return { ...block, subContent: newSubContent };
        }
      }
      return block;
    });
  };

  const setBlocksWithCount = (newBlocksOrUpdater: Block[] | ((prev: Block[]) => Block[])) => {
    setBlocks(prev => {
      const nextBlocks = typeof newBlocksOrUpdater === 'function' 
        ? newBlocksOrUpdater(prev) 
        : newBlocksOrUpdater;
      return recalculateCounts(nextBlocks);
    });
  };

  // --- Actions ---

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newBlocks: Block[] = [];
      Array.from(e.target.files).forEach((item) => {
        const file = item as File;
        if (file.type.startsWith('image/')) {
          const url = URL.createObjectURL(file);
          newBlocks.push({
            id: generateId(),
            type: BlockType.IMAGE,
            content: url,
          });
        }
      });
      setBlocksWithCount(prev => [...prev, ...newBlocks]);
    }
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const addTextRow = () => {
    setBlocksWithCount(prev => [...prev, {
      id: generateId(),
      type: BlockType.TEXT_ROW,
      content: '新文本',
      subContent: '好评0次'
    }]);
  };

  const addTitleRow = () => {
    setBlocks(prev => [...prev, {
      id: generateId(),
      type: BlockType.TITLE,
      content: '1月份网络平台表扬表'
    }]);
  };

  const updateBlock = (id: string, updates: Partial<Block>) => {
    setBlocks(prev => prev.map(b => b.id === id ? { ...b, ...updates } : b));
  };

  const removeBlock = (id: string) => {
    setBlocksWithCount(prev => prev.filter(b => b.id !== id));
  };

  // --- Margin Scrubber Logic ---
  const handleMarginScrub = (e: React.MouseEvent | React.TouchEvent) => {
    // Basic support for touch scrubbing? Maybe later. Mouse for now.
    if ('touches' in e) return; // Skip touch for this specific scrub interaction logic for now
    
    e.preventDefault();
    const startX = e.clientX;
    const startMargin = margin;
    
    const onMouseMove = (me: MouseEvent) => {
        const delta = Math.floor((me.clientX - startX) / 2); // Divide by 2 for smoother control
        const newMargin = Math.max(0, Math.min(50, startMargin + delta));
        setMargin(newMargin);
    };

    const onMouseUp = () => {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        document.body.style.cursor = 'default';
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
    document.body.style.cursor = 'ew-resize';
  };

  // --- Save / Load Project ---

  const getSmartFilename = () => {
    let filename = "layout-project";
    
    // 1. Try Title Block
    const titleBlock = blocks.find(b => b.type === BlockType.TITLE && b.content && b.content.trim().length > 0 && b.content !== '文档标题');
    if (titleBlock) {
      filename = titleBlock.content.trim();
    } else {
      // 2. Try Text Row (Main Content)
      const textRow = blocks.find(b => b.type === BlockType.TEXT_ROW && b.content && b.content.trim().length > 0 && b.content !== '新文本');
      if (textRow) {
          filename = textRow.content.trim();
      }
    }
    // Sanitize filename
    return filename.replace(/[<>:"/\\|?*]/g, '_');
  };

  const handleSaveProject = async () => {
    try {
      const portableBlocks = await Promise.all(blocks.map(async (b) => {
        if (b.type === BlockType.IMAGE) {
          const base64 = await blobUrlToBase64(b.content);
          return { ...b, content: base64 };
        }
        return b;
      }));

      const projectData = {
        version: 1,
        margin,
        blocks: portableBlocks
      };

      const filename = getSmartFilename();
      const blob = new Blob([JSON.stringify(projectData)], { type: "application/json;charset=utf-8" });
      FileSaver.saveAs(blob, `${filename}.json`);
    } catch (error) {
      console.error("Save failed", error);
      alert("保存失败");
    }
  };

  const handleLoadProject = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const json = JSON.parse(event.target?.result as string);
          if (json.blocks) {
            setBlocks(recalculateCounts(json.blocks));
          }
          if (json.margin !== undefined) {
            setMargin(json.margin);
          }
        } catch (err) {
          alert("无法读取项目文件，格式可能错误。");
        }
      };
      reader.readAsText(file);
    }
    if (jsonInputRef.current) jsonInputRef.current.value = '';
  };

  const handleExportDocx = async () => {
    try {
      const filename = getSmartFilename();
      await exportToDocx(blocks, margin, filename);
    } catch (error) {
      console.error("Export failed", error);
      alert("导出失败，请检查控制台。");
    }
  };

  // --- Drag & Drop Logic ---
  
  const handleDragStart = (e: React.DragEvent, index: number) => {
    if (editingBlockId !== null) {
        e.preventDefault();
        return;
    }
    setDraggedBlockIndex(index);
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("application/react-dnd-internal", "true");
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault(); // Necessary to allow dropping

    // External files
    if (e.dataTransfer.types.includes("Files")) {
        e.dataTransfer.dropEffect = "copy";
        setTargetBlockIndex(index); 
        return;
    }

    // Internal Drag
    if (draggedBlockIndex !== null && draggedBlockIndex !== index) {
       setTargetBlockIndex(index);
       e.dataTransfer.dropEffect = "move";
    }
  };

  const handleDragEnd = () => {
    setDraggedBlockIndex(null);
    setTargetBlockIndex(null);
  };

  // Handler for dropping ONTO a specific block
  const handleBlockDrop = (e: React.DragEvent, dropIndex: number) => {
    e.preventDefault();
    e.stopPropagation(); // Stop bubbling to container

    // 1. Handle External File Drop -> Always Insert (Green style logic)
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        const files = Array.from(e.dataTransfer.files).filter((f) => (f as File).type.startsWith('image/')) as File[];
        if (files.length > 0) {
            const newBlocks: Block[] = files.map(file => ({
                id: generateId(),
                type: BlockType.IMAGE,
                content: URL.createObjectURL(file),
            }));

            setBlocksWithCount(prev => {
                const updatedBlocks = [...prev];
                updatedBlocks.splice(dropIndex, 0, ...newBlocks);
                return updatedBlocks;
            });
        }
        setTargetBlockIndex(null);
        return;
    }

    // 2. Handle Internal Drop
    if (draggedBlockIndex !== null && draggedBlockIndex !== dropIndex) {
        const sourceBlock = blocks[draggedBlockIndex];
        const targetBlock = blocks[dropIndex];

        // Logic Check: Swap or Insert?
        // Swap: Source is Image AND Target is Image
        const isSwapOperation = sourceBlock.type === BlockType.IMAGE && targetBlock.type === BlockType.IMAGE;

        setBlocksWithCount(prev => {
            const newBlocks = [...prev];

            if (isSwapOperation) {
                // Swap logic: Exchange positions directly
                newBlocks[draggedBlockIndex] = targetBlock;
                newBlocks[dropIndex] = sourceBlock;
            } else {
                // Insert/Push logic: Standard reorder
                const [movedItem] = newBlocks.splice(draggedBlockIndex, 1);
                
                // Fix: When dragging downwards (e.g., from index 0 to index 5), 
                // removing item at 0 shifts all subsequent items (1->0, 2->1, ... 5->4).
                // The target dropIndex (5) now refers to what was originally at 6, or simply is off by one relative to the new array.
                // Visually, the green line is "Before Target". 
                // If dragged < drop, the target has shifted left by 1. So we must insert at dropIndex - 1.
                let insertIndex = dropIndex;
                if (draggedBlockIndex < dropIndex) {
                    insertIndex -= 1;
                }
                newBlocks.splice(insertIndex, 0, movedItem);
            }
            return newBlocks;
        });
    }

    setDraggedBlockIndex(null);
    setTargetBlockIndex(null);
  };

  // Handler for dropping ON THE CONTAINER (Background) - Append to end
  const handleContainerDrop = (e: React.DragEvent) => {
    e.preventDefault();
    
    // External files -> Append
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        const files = Array.from(e.dataTransfer.files).filter((f) => (f as File).type.startsWith('image/')) as File[];
        if (files.length === 0) return;

        const newBlocks: Block[] = files.map(file => ({
            id: generateId(),
            type: BlockType.IMAGE,
            content: URL.createObjectURL(file),
        }));

        setBlocksWithCount(prev => [...prev, ...newBlocks]);
    } else if (draggedBlockIndex !== null) {
        // Internal drag dropped on whitespace -> move to end
        setBlocksWithCount(prev => {
            const newBlocks = [...prev];
            const [movedItem] = newBlocks.splice(draggedBlockIndex, 1);
            newBlocks.push(movedItem);
            return newBlocks;
        });
    }
    setDraggedBlockIndex(null);
    setTargetBlockIndex(null);
  };

  const handleContainerDragOver = (e: React.DragEvent) => {
      e.preventDefault();
      if (e.target === e.currentTarget) {
         setTargetBlockIndex(null);
      }
  };

  return (
    <div className="min-h-screen bg-[#333333] flex flex-col items-center py-5 font-sans relative">
      <style>{`
        /* Hide Spinner for Number Input */
        input[type=number]::-webkit-inner-spin-button, 
        input[type=number]::-webkit-outer-spin-button { 
          -webkit-appearance: none; 
          margin: 0; 
        }
        input[type=number] {
          -moz-appearance: textfield;
        }
      `}</style>
      
      {/* --- Modern Floating Toolbar (Responsive) --- */}
      <div className="sticky top-4 z-50 mb-8 max-w-[95vw]">
        {/* We use whitespace-nowrap and overflow-auto for horizontal scrolling on mobile */}
        <div className="flex items-center gap-1.5 p-2 bg-white/95 backdrop-blur-md rounded-2xl shadow-2xl border border-gray-100/50 overflow-x-auto no-scrollbar snap-x">
          
          {/* Group 1: File Operations (Colored) */}
          <div className="flex items-center gap-1.5 px-1 shrink-0 snap-start">
             <button 
                onClick={() => jsonInputRef.current?.click()}
                className="flex flex-col items-center justify-center w-12 h-12 gap-1 text-orange-600 bg-orange-50 hover:bg-orange-100 rounded-xl transition-all duration-200 border border-orange-100 group"
                title="导入工程"
              >
                <FolderOpen size={18} className="group-hover:scale-110 transition-transform" />
                <span className="text-[10px] font-medium leading-none">导入</span>
              </button>
              <button 
                onClick={handleSaveProject}
                className="flex flex-col items-center justify-center w-12 h-12 gap-1 text-emerald-600 bg-emerald-50 hover:bg-emerald-100 rounded-xl transition-all duration-200 border border-emerald-100 group"
                title="保存工程"
              >
                <Save size={18} className="group-hover:scale-110 transition-transform" />
                <span className="text-[10px] font-medium leading-none">保存</span>
              </button>
          </div>

          <div className="w-[1px] h-8 bg-gray-200 mx-1 shrink-0"></div>

          {/* Group 2: Insert Tools */}
          <div className="flex items-center gap-1.5 px-1 shrink-0 snap-start">
            <button 
              onClick={() => fileInputRef.current?.click()}
              className="flex flex-col items-center justify-center w-12 h-12 gap-1 text-gray-700 hover:text-black hover:bg-gray-100 rounded-xl transition-all duration-200 group"
              title="添加图片"
            >
              <ImageIcon size={20} className="group-hover:scale-110 transition-transform" />
              <span className="text-[10px] font-medium leading-none text-gray-500 group-hover:text-black">图片</span>
            </button>
            <button 
              onClick={addTextRow}
              className="flex flex-col items-center justify-center w-12 h-12 gap-1 text-gray-700 hover:text-black hover:bg-gray-100 rounded-xl transition-all duration-200 group"
              title="添加文本行"
            >
              <Type size={20} className="group-hover:scale-110 transition-transform" />
              <span className="text-[10px] font-medium leading-none text-gray-500 group-hover:text-black">文本</span>
            </button>
            <button 
              onClick={addTitleRow}
              className="flex flex-col items-center justify-center w-12 h-12 gap-1 text-gray-700 hover:text-black hover:bg-gray-100 rounded-xl transition-all duration-200 group"
              title="添加标题"
            >
              <Heading size={20} className="group-hover:scale-110 transition-transform" />
              <span className="text-[10px] font-medium leading-none text-gray-500 group-hover:text-black">标题</span>
            </button>
          </div>

          <div className="w-[1px] h-8 bg-gray-200 mx-1 shrink-0"></div>

          {/* Group 3: Margin Control (Clean & Scrubbable) */}
          <div className="flex flex-col justify-center px-2 min-w-[60px] shrink-0 snap-start">
             <div 
                className="flex items-center gap-1 text-gray-400 hover:text-blue-500 cursor-ew-resize select-none transition-colors group py-1"
                onMouseDown={handleMarginScrub}
                title="按住左右拖动调整"
             >
                <MoveHorizontal size={12} />
                <span className="text-[10px] font-bold uppercase tracking-wider">间距</span>
             </div>
             <input 
                type="number" 
                min="0" 
                max="50"
                value={margin}
                onChange={(e) => setMargin(Number(e.target.value))}
                className="w-full text-center text-sm font-semibold text-gray-700 bg-transparent border-none focus:ring-0 p-0 m-0 hover:text-black appearance-none"
             />
          </div>

          <div className="w-[1px] h-8 bg-gray-200 mx-1 shrink-0"></div>

          {/* Group 4: Export & Help */}
          <div className="flex items-center gap-2 px-1 shrink-0 snap-start">
             <button 
                onClick={handleExportDocx}
                className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white pl-4 pr-5 py-2.5 rounded-xl text-sm font-semibold shadow-md shadow-blue-200 transition-all active:scale-95 whitespace-nowrap"
            >
                <FileDown size={18} />
                导出 Word
            </button>
            
            <button 
                onClick={() => { setShowHelp(true); setHelpTab('guide'); }}
                className="flex items-center justify-center w-10 h-10 text-gray-400 hover:text-blue-500 hover:bg-blue-50 rounded-full transition-all shrink-0"
                title="使用帮助"
            >
                <HelpCircle size={20} />
            </button>
          </div>
        </div>
      </div>

      {/* --- Help / Updates Modal --- */}
      {showHelp && (
        <div 
            className="fixed inset-0 z-[100] flex items-center justify-center bg-black/30 backdrop-blur-sm p-4 animate-in fade-in duration-200"
            onClick={() => setShowHelp(false)}
        >
            <div 
                className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden border border-gray-100 animate-in zoom-in-95 duration-200 max-h-[80vh] overflow-y-auto"
                onClick={e => e.stopPropagation()}
            >
                {/* Modal Header & Tabs */}
                <div className="flex flex-col border-b border-gray-100 bg-gray-50/50">
                    <div className="flex items-center justify-between p-4 pb-2">
                        <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                           <div className="bg-blue-100 p-1.5 rounded-lg text-blue-600">
                                {helpTab === 'guide' ? <BookOpen size={18} /> : <Sparkles size={18} />}
                           </div>
                           {helpTab === 'guide' ? '操作指南' : '更新日志'}
                        </h3>
                        <button onClick={() => setShowHelp(false)} className="text-gray-400 hover:text-black hover:bg-gray-100 p-1 rounded-full transition-colors">
                            <X size={20} />
                        </button>
                    </div>

                    {/* Tabs */}
                    <div className="flex px-4 gap-6 text-sm font-medium">
                        <button 
                            onClick={() => setHelpTab('guide')}
                            className={`pb-2 border-b-2 transition-colors ${helpTab === 'guide' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'}`}
                        >
                            操作指南
                        </button>
                        <button 
                            onClick={() => setHelpTab('updates')}
                            className={`pb-2 border-b-2 transition-colors ${helpTab === 'updates' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'}`}
                        >
                            最近更新
                        </button>
                    </div>
                </div>
                
                {/* Modal Body */}
                <div className="p-6 min-h-[300px]">
                    
                    {/* --- TAB: GUIDE --- */}
                    {helpTab === 'guide' && (
                        <div className="space-y-6 animate-in slide-in-from-right-4 duration-300">
                            {/* Tip 1: Color Codes */}
                            <div className="flex gap-4">
                                <div className="mt-1">
                                    <RefreshCw className="text-blue-500" size={24} />
                                </div>
                                <div>
                                    <h4 className="text-base font-bold text-gray-800 mb-1">拖拽颜色指示</h4>
                                    <div className="space-y-2 text-sm text-gray-600 leading-relaxed">
                                        <p className="flex items-center gap-2">
                                            <span className="w-3 h-3 rounded-full bg-blue-500 inline-block shadow-sm"></span>
                                            <strong className="text-blue-600">蓝色边框：</strong> 
                                            交换位置 (图片 ⇄ 图片)。
                                        </p>
                                        <p className="flex items-center gap-2">
                                            <span className="w-3 h-3 rounded-full bg-green-500 inline-block shadow-sm"></span>
                                            <strong className="text-green-600">绿色横线：</strong> 
                                            插入/挤压位置。
                                        </p>
                                    </div>
                                </div>
                            </div>

                            <div className="h-[1px] bg-gray-100 w-full"></div>

                            {/* Tip 2: Crop */}
                            <div className="flex gap-4">
                                <div className="mt-1">
                                    <Scissors className="text-orange-500" size={24} />
                                </div>
                                <div>
                                    <h4 className="text-base font-bold text-gray-800 mb-1">图片裁切</h4>
                                    <p className="text-sm text-gray-600 leading-relaxed">
                                        鼠标移动到图片<strong className="text-gray-900 bg-gray-100 px-1 rounded">底部边缘</strong>，按住并向上拖动，可以裁切高度（用于去除无关底部）。
                                    </p>
                                </div>
                            </div>

                            <div className="h-[1px] bg-gray-100 w-full"></div>

                            {/* Tip 3: Margin */}
                            <div className="flex gap-4">
                                <div className="mt-1">
                                    <ArrowDownToLine className="text-purple-500" size={24} />
                                </div>
                                <div>
                                    <h4 className="text-base font-bold text-gray-800 mb-1">快速调整间距</h4>
                                    <p className="text-sm text-gray-600 leading-relaxed">
                                        在工具栏上，按住 <strong className="bg-gray-100 px-1 py-0.5 rounded text-gray-700 text-xs uppercase">↔ 间距</strong> 标签并左右拖动，或直接输入数字。
                                    </p>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* --- TAB: UPDATES --- */}
                    {helpTab === 'updates' && (
                        <div className="space-y-6 animate-in slide-in-from-right-4 duration-300">
                             {/* Item 0 (New) */}
                             <div className="flex gap-4">
                                <div className="flex flex-col items-center">
                                    <div className="w-8 h-8 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0">
                                        <Scissors size={16} />
                                    </div>
                                    <div className="w-[2px] h-full bg-emerald-50 mt-2"></div>
                                </div>
                                <div className="pb-4">
                                    <h4 className="text-base font-bold text-gray-800">裁切体验优化</h4>
                                    <p className="text-sm text-gray-500 mb-1">2026.01.07</p>
                                    <p className="text-sm text-gray-600 leading-relaxed">
                                        优化了图片裁切时的页面布局稳定性。现在向上拖动裁切时，A4纸张高度不会实时塌缩，防止页面跳动，操作更流畅。
                                    </p>
                                </div>
                             </div>

                             {/* Item 1 */}
                             <div className="flex gap-4">
                                <div className="flex flex-col items-center">
                                    <div className="w-8 h-8 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center shrink-0">
                                        <MousePointerClick size={16} />
                                    </div>
                                    <div className="w-[2px] h-full bg-blue-50 mt-2"></div>
                                </div>
                                <div className="pb-4">
                                    <h4 className="text-base font-bold text-gray-800">拖拽逻辑优化</h4>
                                    <p className="text-sm text-gray-500 mb-1">2026.01.07</p>
                                    <p className="text-sm text-gray-600 leading-relaxed">
                                        区分了“交换”与“插入”操作。现在拖拽图片到另一张图片时会显示<span className="text-blue-600 font-medium">蓝色框</span>（交换），拖到文本时显示<span className="text-green-600 font-medium">绿色线</span>（插入）。
                                    </p>
                                </div>
                             </div>

                             {/* Item 2 */}
                             <div className="flex gap-4">
                                <div className="flex flex-col items-center">
                                    <div className="w-8 h-8 rounded-full bg-orange-50 text-orange-600 flex items-center justify-center shrink-0">
                                        <Sparkles size={16} />
                                    </div>
                                    <div className="w-[2px] h-full bg-orange-50 mt-2"></div>
                                </div>
                                <div className="pb-4">
                                    <h4 className="text-base font-bold text-gray-800">智能统计与命名</h4>
                                    <p className="text-sm text-gray-500 mb-1">2026.01.07</p>
                                    <ul className="text-sm text-gray-600 leading-relaxed list-disc list-inside space-y-1">
                                        <li>自动统计文本行下方的图片数量，并更新“好评x次”。</li>
                                        <li>保存工程与导出文档时，自动使用标题或首行文本作为文件名。</li>
                                    </ul>
                                </div>
                             </div>

                             {/* Item 3 */}
                             <div className="flex gap-4">
                                <div className="flex flex-col items-center">
                                    <div className="w-8 h-8 rounded-full bg-purple-50 text-purple-600 flex items-center justify-center shrink-0">
                                        <MoveHorizontal size={16} />
                                    </div>
                                </div>
                                <div>
                                    <h4 className="text-base font-bold text-gray-800">全新 UI 设计</h4>
                                    <p className="text-sm text-gray-500 mb-1">2026.01.07</p>
                                    <p className="text-sm text-gray-600 leading-relaxed">
                                        采用了悬浮式工具栏设计，优化了间距调节控件，整体视觉更现代。
                                    </p>
                                </div>
                             </div>
                        </div>
                    )}
                </div>

                {/* Footer */}
                <div className="p-4 bg-gray-50 border-t border-gray-100 text-center">
                    <button 
                        onClick={() => setShowHelp(false)}
                        className="text-sm font-medium text-blue-600 hover:text-blue-700 hover:underline"
                    >
                        关闭窗口
                    </button>
                </div>
            </div>
        </div>
      )}

      <input 
          type="file" 
          ref={fileInputRef} 
          hidden 
          multiple 
          accept="image/*" 
          onChange={handleImageUpload} 
        />
      <input 
          type="file" 
          ref={jsonInputRef} 
          hidden 
          accept=".json" 
          onChange={handleLoadProject} 
      />

      {/* --- Page Container --- */}
      <div id="page-container" className="flex flex-col items-center w-full overflow-y-auto pb-10 px-2 sm:px-0">
        <div 
            ref={pageContainerRef}
            className="a4-page bg-white p-[10mm] box-border flex flex-wrap content-start gap-[10px] shadow-2xl relative transition-all origin-top"
            onDrop={handleContainerDrop}
            onDragOver={handleContainerDragOver}
            style={{ 
                width: '210mm', 
                minHeight: '297mm', // A4 Height
                transform: `scale(${scale})`,
                marginBottom: `calc(297mm * ${scale - 1})` // Adjust margin to remove ghost space from scaling
            }}
        >
            {blocks.length === 0 && (
                <div className="absolute inset-0 m-[10mm] flex flex-col items-center justify-center text-gray-300 border-2 border-dashed border-gray-100 rounded-lg pointer-events-none">
                    <p className="font-medium">画布为空</p>
                    <p className="text-sm">拖入图片 或 点击工具栏添加</p>
                </div>
            )}

            {blocks.map((block, index) => {
                // Determine Visual State
                const isSource = draggedBlockIndex === index;
                const isTarget = targetBlockIndex === index && !isSource;
                
                let visualClass = '';
                
                if (isTarget) {
                    // Decide Swap vs Insert Visuals
                    // Case 1: External File Drag -> Always Insert (Green)
                    // We don't have easy access to 'isExternal' here without checking a global or state, 
                    // BUT draggedBlockIndex is null during external drag.
                    const isExternalDrag = draggedBlockIndex === null;

                    if (isExternalDrag) {
                         visualClass = 'border-t-[4px] border-green-500 pt-0';
                    } else {
                        // Internal Drag
                        const sourceBlock = blocks[draggedBlockIndex];
                        const targetBlock = blocks[index];
                        
                        const isSwapMode = sourceBlock.type === BlockType.IMAGE && targetBlock.type === BlockType.IMAGE;

                        if (isSwapMode) {
                            // Swap Visual: Blue Border Box
                            visualClass = 'border-2 border-blue-500 rounded-sm scale-[1.02] z-10';
                        } else {
                            // Insert Visual: Green Top Border
                            visualClass = 'border-t-[4px] border-green-500 pt-0';
                        }
                    }
                }

                return (
                <div
                    key={block.id}
                    data-block-index={index} 
                    draggable={editingBlockId !== block.id}
                    onDragStart={(e) => handleDragStart(e, index)}
                    onDragOver={(e) => handleDragOver(e, index)}
                    onDrop={(e) => handleBlockDrop(e, index)}
                    onDragEnd={handleDragEnd}
                    className={`
                        ${block.type === BlockType.IMAGE ? 'w-[calc(50%-5px)]' : 'w-full'}
                        transition-all duration-100 ease-out cursor-move relative
                        ${isSource ? 'opacity-30 scale-95' : 'opacity-100'}
                        ${visualClass}
                    `}
                    style={{ marginBottom: `${margin}px` }}
                >
                    <BlockRenderer 
                        block={block} 
                        onUpdate={updateBlock} 
                        onRemove={removeBlock}
                        isDragging={isSource}
                        onEditStart={() => setEditingBlockId(block.id)}
                        onEditEnd={() => setEditingBlockId(null)}
                    />
                </div>
            )})}
        </div>
      </div>
    </div>
  );
}

export default App;