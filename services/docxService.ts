import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  ImageRun, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType,
  AlignmentType,
  BorderStyle,
  TableLayoutType
} from "docx";
import FileSaver from "file-saver";
import { Block, BlockType } from "../types";

// Constants for A4 Layout in Twips (1/1440 inch)
// 1mm approx 56.69 twips
const PAGE_MARGIN = 567; // 10mm (1cm)
const PAGE_WIDTH = 11906; // 210mm (A4 width)
const CONTENT_WIDTH = PAGE_WIDTH - (PAGE_MARGIN * 2); // 10772
const COLUMN_WIDTH = Math.floor(CONTENT_WIDTH / 2); // 5386

// Helper to determine image type from blob type
const getImageType = (mimeType: string): "png" | "jpg" | "gif" | "bmp" => {
    if (mimeType.includes("png")) return "png";
    if (mimeType.includes("gif")) return "gif";
    if (mimeType.includes("bmp")) return "bmp";
    return "jpg";
};

// Helper to convert base64/url to ArrayBuffer for docx
// Now supports cropping via Canvas if a cropHeight is provided
const TRANSPARENT_PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0,
  0, 0, 11, 73, 68, 65, 84, 8, 215, 99, 96, 0, 2, 0, 0, 5, 0,
  1, 226, 38, 5, 155, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130
]).buffer;

const processImage = async (url: string, cropHeight?: number): Promise<{ data: ArrayBuffer, type: "png" | "jpg" | "gif" | "bmp" }> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    
    const resolveFallback = () => {
        resolve({ data: TRANSPARENT_PNG, type: "png" });
    };

    const resolveOriginal = () => {
        if (url.startsWith('data:')) {
            // Parse data URI manually to avoid fetch issues on large strings
            try {
                const arr = url.split(',');
                const bstr = atob(arr[1]);
                let n = bstr.length;
                const u8arr = new Uint8Array(n);
                while(n--){
                    u8arr[n] = bstr.charCodeAt(n);
                }
                const mime = arr[0].match(/:(.*?);/)?.[1] || 'image/png';
                resolve({ data: u8arr.buffer, type: getImageType(mime) });
            } catch (e) {
                resolveFallback();
            }
            return;
        }

        fetch(url)
            .then(r => {
                if (!r.ok) throw new Error('Network response was not ok');
                return r.blob();
            })
            .then(async blob => {
                resolve({ 
                    data: await blob.arrayBuffer(), 
                    type: getImageType(blob.type) 
                });
            })
            .catch((err) => {
                console.warn("Failed to fetch image:", url, err);
                resolveFallback();
            });
    };

    img.onload = () => {
      // Logic: 
      // The web editor displays images at approx 354px width (Half of A4 210mm - padding).
      // We need to calculate the crop ratio based on that to match "what you see".
      const REFERENCE_DISPLAY_WIDTH = 354; 
      
      if (cropHeight) {
        // Calculate expected height if image was full width (scale logic)
        // If cropHeight >= expectedHeight, we don't need to crop (or user expanded it fully)
        // Actually, if cropHeight is present, we should trust it, but we need to map 
        // screen pixels to image pixels.
        
        const scale = img.naturalWidth / REFERENCE_DISPLAY_WIDTH;
        const targetIntrinsicHeight = cropHeight * scale;

        // If target height covers the whole image, just return original
        if (targetIntrinsicHeight >= img.naturalHeight - 1) { // epsilon
             resolveOriginal();
             return;
        }

        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        
        canvas.width = img.naturalWidth;
        canvas.height = targetIntrinsicHeight;

        if (ctx) {
          // Draw image top-aligned. The bottom part is cut off.
          ctx.drawImage(img, 0, 0);
          
          canvas.toBlob((blob) => {
             if (blob) {
                blob.arrayBuffer().then(buffer => {
                    resolve({ data: buffer, type: "png" });
                });
             }
             else {
                resolveFallback();
             }
          }, 'image/png');
        } else {
           resolveOriginal();
        }
      } else {
        resolveOriginal();
      }
    };

    img.onerror = () => {
      resolveOriginal();
    };
    
    img.src = url;
  });
};

// Helper to get Aspect Ratio (H/W) for Docx scaling
const getImageAspectRatio = (url: string): Promise<number> => {
    return new Promise((resolve) => {
        const img = new Image();
        img.src = url;
        img.onload = () => {
            resolve(img.naturalHeight / img.naturalWidth);
        };
        img.onerror = () => resolve(1);
    });
};

export const exportToDocx = async (blocks: Block[], margin: number, filename: string) => {
  const docChildren: (Paragraph | Table)[] = [];
  
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];

    if (block.type === BlockType.TITLE) {
      docChildren.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200 },
          children: [
            new TextRun({
              text: block.content,
              bold: true,
              size: 32, // 16pt
              color: "000000"
            }),
          ],
        })
      );
    } 
    
    else if (block.type === BlockType.TEXT_ROW) {
      // Export Text Row with Main content and Sub content
      const runs = [
          new TextRun({
              text: block.content,
              bold: true,
              size: 24, // 12pt
              color: "333333"
          })
      ];

      if (block.subContent) {
          // Add a tab or space separator
          runs.push(new TextRun({ text: "    " })); 
          runs.push(new TextRun({
              text: block.subContent,
              size: 20, // 10pt
              color: "666666"
          }));
      }

      docChildren.push(
        new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 100, after: 100 },
            children: runs
        })
      );
    } 
    
    else if (block.type === BlockType.IMAGE) {
      const nextBlock = blocks[i + 1];
      const isNextImage = nextBlock && nextBlock.type === BlockType.IMAGE;

      const img1 = await processImage(block.content, block.height);
      
      if (isNextImage) {
        const img2 = await processImage(nextBlock.content, nextBlock.height);

        const table = new Table({
            layout: TableLayoutType.FIXED, 
            alignment: AlignmentType.CENTER, 
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            columnWidths: [COLUMN_WIDTH, COLUMN_WIDTH], // Explicitly set column widths
            borders: {
                top: { style: BorderStyle.NIL, size: 0, color: "auto" },
                bottom: { style: BorderStyle.NIL, size: 0, color: "auto" },
                left: { style: BorderStyle.NIL, size: 0, color: "auto" },
                right: { style: BorderStyle.NIL, size: 0, color: "auto" },
                insideVertical: { style: BorderStyle.NIL, size: 0, color: "auto" },
                insideHorizontal: { style: BorderStyle.NIL, size: 0, color: "auto" },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: { size: COLUMN_WIDTH, type: WidthType.DXA },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.CENTER, 
                                    children: [
                                        new ImageRun({
                                            data: img1.data,
                                            transformation: { width: 300, height: block.height ? (300 * block.height / 354) : (300 * (await getImageAspectRatio(block.content))) }, 
                                            type: img1.type,
                                        })
                                    ]
                                })
                            ]
                        }),
                        new TableCell({
                            width: { size: COLUMN_WIDTH, type: WidthType.DXA },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.CENTER, 
                                    children: [
                                        new ImageRun({
                                            data: img2.data,
                                            transformation: { width: 300, height: nextBlock.height ? (300 * nextBlock.height / 354) : (300 * (await getImageAspectRatio(nextBlock.content))) },
                                            type: img2.type,
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                })
            ]
        });
        docChildren.push(table);
        i++; // Skip next
      } else {
        // Single Image
        const table = new Table({
            layout: TableLayoutType.FIXED, 
            alignment: AlignmentType.CENTER,
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            columnWidths: [COLUMN_WIDTH, COLUMN_WIDTH],
            borders: {
                top: { style: BorderStyle.NIL, size: 0, color: "auto" },
                bottom: { style: BorderStyle.NIL, size: 0, color: "auto" },
                left: { style: BorderStyle.NIL, size: 0, color: "auto" },
                right: { style: BorderStyle.NIL, size: 0, color: "auto" },
                insideVertical: { style: BorderStyle.NIL, size: 0, color: "auto" },
                insideHorizontal: { style: BorderStyle.NIL, size: 0, color: "auto" },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: { size: COLUMN_WIDTH, type: WidthType.DXA },
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.CENTER, 
                                    children: [
                                        new ImageRun({
                                            data: img1.data,
                                            transformation: { width: 300, height: block.height ? (300 * block.height / 354) : (300 * (await getImageAspectRatio(block.content))) }, 
                                            type: img1.type,
                                        })
                                    ]
                                })
                            ]
                        }),
                        new TableCell({ 
                            width: { size: COLUMN_WIDTH, type: WidthType.DXA },
                            children: [new Paragraph({})] // Empty paragraph to preserve cell structure
                        })
                    ]
                })
            ]
        });
        docChildren.push(table);
      }
      
      docChildren.push(new Paragraph({ spacing: { after: margin * 10 } })); 
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {
            page: {
                margin: {
                    top: PAGE_MARGIN,
                    bottom: PAGE_MARGIN,
                    left: PAGE_MARGIN,
                    right: PAGE_MARGIN
                }
            }
        },
        children: docChildren,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  FileSaver.saveAs(blob, `${filename}.docx`);
};
