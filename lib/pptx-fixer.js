const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

/**
 * PPTX Post-Processor: Fixes corruption issues in generated PPTX files
 * 
 * ROOT CAUSES FIXED:
 * 1. Empty name attributes in p:cNvPr elements
 * 2. Empty a:ln elements with no attributes
 * 3. Zero dimensions in group shape transforms
 * 4. Conflicting autofit settings (normAutofit + spAutoFit)
 * 5. Very small dimension values that PowerPoint rejects
 * 
 * These issues originate from PptxGenJS 3.12.0 and cause PowerPoint
 * to mark files as corrupted and require repair.
 */
class PPTXFixer {
    constructor() {
        this.fixCount = 0;
    }

    /**
     * Fix a PPTX file by correcting XML corruption issues
     * @param {string} filePath - Path to PPTX file to fix
     */
    async fixPPTX(filePath) {
        try {
            console.log(`[PPTX Fixer] Processing: ${filePath}`);
            this.fixCount = 0;

            // Load the PPTX as a ZIP file
            const zip = new AdmZip(filePath);
            const zipEntries = zip.getEntries();

            // Process all XML files
            for (const entry of zipEntries) {
                if (entry.entryName.endsWith('.xml') || entry.entryName.endsWith('.rels')) {
                    const content = entry.getData().toString('utf8');
                    const fixed = this.fixXML(content, entry.entryName);
                    
                    if (fixed !== content) {
                        // Update the entry with fixed content
                        entry.setData(Buffer.from(fixed, 'utf8'));
                    }
                }
            }

            // Save the fixed PPTX
            if (this.fixCount > 0) {
                // Create backup
                const backup = filePath + '.backup';
                fs.copyFileSync(filePath, backup);
                
                // Save fixed version
                zip.writeZip(filePath);
                console.log(`[PPTX Fixer] Applied ${this.fixCount} fixes, backup saved to ${backup}`);
                
                // Remove backup if successful
                fs.unlinkSync(backup);
            } else {
                console.log(`[PPTX Fixer] No issues found`);
            }

            return {
                success: true,
                fixCount: this.fixCount
            };
        } catch (error) {
            console.error(`[PPTX Fixer] Error: ${error.message}`);
            throw error;
        }
    }

    /**
     * Fix XML content
     * @param {string} xml - XML content
     * @param {string} filename - File name for logging
     */
    fixXML(xml, filename) {
        let fixed = xml;

        // FIX 1: Empty name attributes in p:cNvPr
        // Pattern: <p:cNvPr id="X" name=""/>
        // Fix: <p:cNvPr id="X" name="Shape X"/>
        fixed = fixed.replace(
            /<p:cNvPr\s+id="(\d+)"\s+name=""\s*\/>/g,
            (match, id) => {
                this.fixCount++;
                return `<p:cNvPr id="${id}" name="Shape ${id}"/>`;
            }
        );
        
        // Also handle non-self-closing variant
        fixed = fixed.replace(
            /<p:cNvPr\s+id="(\d+)"\s+name="">([^<]*)<\/p:cNvPr>/g,
            (match, id, content) => {
                this.fixCount++;
                return `<p:cNvPr id="${id}" name="Shape ${id}">${content}</p:cNvPr>`;
            }
        );

        // FIX 2: Empty a:ln elements
        // Pattern: <a:ln></a:ln> or <a:ln/>
        // Fix: Remove them entirely (no border)
        const beforeLnFix = fixed;
        fixed = fixed.replace(/<a:ln\s*><\/a:ln>/g, () => {
            this.fixCount++;
            return '';
        });
        fixed = fixed.replace(/<a:ln\s*\/>/g, () => {
            // Only count if we actually removed something
            if (fixed !== beforeLnFix) return '';
            this.fixCount++;
            return '';
        });

        // FIX 3: Zero dimensions in transforms
        // Pattern: <a:ext cx="0" cy="0"/>
        // Fix: Use minimum valid dimensions (1 EMU)
        fixed = fixed.replace(
            /<a:ext\s+cx="0"\s+cy="0"\s*\/>/g,
            () => {
                this.fixCount++;
                return '<a:ext cx="1" cy="1"/>';
            }
        );
        
        // Also fix individual zero dimensions
        fixed = fixed.replace(
            /<a:ext\s+cx="0"\s+cy="(\d+)"\s*\/>/g,
            (match, cy) => {
                this.fixCount++;
                return `<a:ext cx="1" cy="${cy}"/>`;
            }
        );
        fixed = fixed.replace(
            /<a:ext\s+cx="(\d+)"\s+cy="0"\s*\/>/g,
            (match, cx) => {
                this.fixCount++;
                return `<a:ext cx="${cx}" cy="1"/>`;
            }
        );

        // FIX 4: Conflicting autofit settings
        // Pattern: <a:bodyPr...><a:normAutofit/><a:spAutoFit/></a:bodyPr>
        // Fix: Remove spAutoFit, keep normAutofit
        fixed = fixed.replace(
            /(<a:bodyPr[^>]*>)((?:(?!<\/a:bodyPr>).)*?)<a:normAutofit\s*\/>((?:(?!<\/a:bodyPr>).)*?)<a:spAutoFit\s*\/>/gi,
            (match, opening, before, after) => {
                this.fixCount++;
                // Keep normAutofit, remove spAutoFit
                return `${opening}${before}<a:normAutofit/>${after}`;
            }
        );

        // FIX 5: Very small dimensions (less than 10000 EMUs)
        // These can cause rendering issues in PowerPoint
        // Pattern: cy="XXXX" where XXXX < 10000
        fixed = fixed.replace(
            /cy="(\d{1,4})"/g,
            (match, value) => {
                const num = parseInt(value);
                if (num > 0 && num < 10000) {
                    this.fixCount++;
                    return 'cy="10000"'; // Minimum 0.14 inches
                }
                return match;
            }
        );

        // FIX 6: Empty or invalid charset attributes
        // Pattern: charset="-122" or charset="-120"
        // Fix: charset="0" (system default)
        fixed = fixed.replace(
            /charset="-?\d+"/g,
            (match) => {
                const val = match.match(/-?\d+/)[0];
                if (parseInt(val) < 0) {
                    this.fixCount++;
                    return 'charset="0"';
                }
                return match;
            }
        );

        return fixed;
    }
}

/**
 * Convenience function to fix a PPTX file
 * @param {string} filePath - Path to PPTX file
 */
async function fixPPTX(filePath) {
    const fixer = new PPTXFixer();
    return await fixer.fixPPTX(filePath);
}

module.exports = {
    PPTXFixer,
    fixPPTX
};
